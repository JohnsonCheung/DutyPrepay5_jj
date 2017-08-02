Attribute VB_Name = "xFmt"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xFmt"
Type tFmtWsOL
    Ws As Worksheet
    RnoFm As Long
    RnoTo As Long
    NLvl As Byte
    ColEnd As String
    AyLvlCols() As String ' Element is in "B:C", "D" format
    RnoColr As Long ' Rno to be use for column: RnoColr & AyLvlCol() will be found
End Type
Function Fmt_WsOutLine(p As tFmtWsOL) As Boolean
'Aim: Color, Outline & Border by {p}
'     Assume Col A is OutLine Level#.
Const cSub$ = "Fmt_WsOutLine"
On Error GoTo R
'Find mAyColr() by p.NLvl & p.AyLvlCols
ReDim mAyColr&(p.NLvl - 1)
Dim iLvl%
For iLvl = 0 To p.NLvl - 1
    Dim mA$: mA = p.AyLvlCols(iLvl)
    Dim mP%: mP% = InStr(mA, ":")
    If mP > 0 Then mA = Left(mA, mP - 1)
    Dim mRge As Range: Set mRge = p.Ws.Cells(p.RnoColr, mA)
    mAyColr(iLvl) = mRge.Interior.Color
Next
'
Dim iRno&
Dim mLvlLas%: mLvlLas = 0
With p.Ws
    For iRno = p.RnoFm To p.RnoTo
        'set Outline
        Dim mOutLine As Byte: mOutLine = .Cells(iRno, 1).Value
        If mOutLine > 1 Then
            Set mRge = .Rows(iRno)
            mRge.EntireRow.OutlineLevel = mOutLine
        End If
    
        If mOutLine = mLvlLas Then GoTo Nx
        mLvlLas = mOutLine
        
        'Find 3 Ranges
        Dim mRgeBorder As Range, mRgeColrRow As Range, mRgeColrCols As Range
        Dim mLvlCur%: mLvlCur = .Cells(iRno, 1).Value
        Dim mLvlCols$: mLvlCols = p.AyLvlCols(mLvlCur - 1)
        Dim mColr&: mColr = mAyColr(mLvlCur - 1)
        If Fmt_WsOutLine_Fnd2Rge(mRgeBorder, mRgeColrCols, p.Ws, iRno, p.RnoTo, mLvlCols, p.ColEnd) Then ss.A 1: GoTo E
        'Fmt the 2 ranges
        mRgeBorder.BorderAround XlLineStyle.xlContinuous, XlBorderWeight.xlMedium
        mRgeBorder.Interior.Color = mColr
      
        If mOutLine <> p.NLvl Then
            Dim mSq As tSq: If xCv.Cv_Rge2Sq(mSq, mRgeColrCols) Then ss.A 2: GoTo E
            If mSq.r2 > mSq.r1 Then
                mSq.r1 = mSq.r1 + 1
                If xCv.Cv_Sq2Rge(mRgeColrCols, p.Ws, mSq) Then ss.A 3: GoTo E
                mRgeColrCols.Font.Color = mColr
            End If
        End If
Nx: Next
End With
Exit Function
R: ss.R
E: Fmt_WsOutLine = True: ss.B cSub, cMod
End Function
Private Function Fmt_WsOutLine_Fnd2Rge(ByRef oRgeBorder As Range, ByRef oRgeColrCols _
    , pWs As Worksheet, pRno&, pRnoTo&, pLvlCols$, pColEnd$) As Boolean
'Aim: Find oRge by p*
'     Assume the Lvl Col is at column A
'     oRgeBorder:   The whole range should be border arround.  Range of Row = Current row to [mRnoEnd],  Range of Col = first col of {pLvlCols} up to pColEnd$
'                   [mRnoEnd] = some row below of lvl =< pLvl or {pRnoTo}
'     pRno: current rno
'     pRnoTo last row
'     pLvlCols in "C" or "C:D" format
'     pColEnd  in "CC" format
Const cSub$ = "Fmt_WsOutLine_Fnd2Rge"
On Error GoTo R
'Find mColBeg
Dim mLvlColBeg$, mLvlColEnd$
Dim mP%: mP = InStr(pLvlCols, ":")
If mP = 0 Then
    mLvlColBeg = pLvlCols
    mLvlColEnd = pLvlCols
Else
    mLvlColBeg = Left(pLvlCols, mP - 1)
    mLvlColEnd = mID(pLvlCols, mP + 1)
End If
'Find mRnoEnd: a Row just before a the row with Lvl =< CurLvl or the pRnoTo
Dim mRnoEnd&
Dim mLvlCur%: mLvlCur = pWs.Cells(pRno, 1).Value
Dim mFound As Boolean: mFound = False
Dim mBrk As Boolean: mBrk = False
For mRnoEnd = pRno + 1 To pRnoTo
    Dim iLvl%: iLvl = pWs.Cells(mRnoEnd, 1).Value
    If mLvlCur = iLvl Then
        If mBrk Then mRnoEnd = mRnoEnd - 1: mFound = True: Exit For
        GoTo Nx
    End If
    mBrk = True
    If iLvl <= mLvlCur Then mRnoEnd = mRnoEnd - 1: mFound = True: Exit For
Nx:
Next
If Not mFound Then mRnoEnd = pRnoTo
'Set oRgeBorder
Set oRgeBorder = pWs.Range(mLvlColBeg & pRno & ":" & pColEnd & mRnoEnd)
Set oRgeColrCols = pWs.Range(mLvlColBeg & pRno & ":" & mLvlColEnd & mRnoEnd)
Exit Function
R: ss.R
E: Fmt_WsOutLine_Fnd2Rge = True: ss.B cSub, cMod
End Function
Function Fmt_Str_Repeat_ByLm_IntoAy(oAys$(), pFmtStr$, pLm_SemiSep$) As Boolean
Const cSub$ = "Fmt_Str_Repeat_ByLm_IntoAy"
Dim mAm() As tMap: mAm = jj.Get_Am_ByLm(pLm_SemiSep, , ";")
Fmt_Str_Repeat_ByLm_IntoAy = Fmt_Str_Repeat_ByAm_IntoAy(oAys, pFmtStr, mAm)
Exit Function
E: Fmt_Str_Repeat_ByLm_IntoAy = True: ss.B cSub, cMod, "pFmtStr,pLm_SemiSep", pFmtStr, pLm_SemiSep
End Function
Function Fmt_Str_Repeat_ByAm_IntoAy(oAys$(), pFmtStr$, pAm() As tMap) As Boolean
'Aim: Find oAys$() by pFmtStr, pLm
'     pFmtStr: contains ....{M1}...{M2}...{M3}...
'     pLm    : contains M1=xx;M2=xx,xx,xx;M3=xx,xx
'     If there is no multiple element in Mx,  {oAys} will return 1 element
'     If there is multiple element in Mx,     {oAys} will return multiple of # of element in each array element in mAp
'                                             example, M2 has 3 elements & M3 has 3 elements
'                                             oAys will return 6 elements
Const cSub$ = "Fmt_Str_Repeat_ByAm_IntoAy"
Dim mS$: mS = pFmtStr
Dim J%, mMacro$
Dim N%: N = jj.Siz_Am(pAm)
For J = 0 To N - 1
    With pAm(J)
        If InStr(.F2, ",") = 0 Then
            mS = Replace(mS, "{" & .F1 & "}", .F2)
        End If
    End With
Next

ReDim mAyFmtStr$(0): mAyFmtStr(0) = mS
Clr_Ays oAys
For J = 0 To N - 1
    With pAm(J)
        If InStr(.F2, ",") <> 0 Then
            mMacro = "{" & .F1 & "}"
            Dim mAys$(): mAys = Split(.F2, ",")
            Dim I%, mAy1$(), mAy$()
            Clr_Ays mAy1
            For I = 0 To jj.Siz_Ay(mAyFmtStr) - 1
                If jj.Fmt_Str_Repeat_Ay_IntoAy(mAy, mAyFmtStr(I), mAys, , mMacro) Then ss.A 1: GoTo E
                If jj.Add_AyAtEnd(mAy1, mAy, True) Then ss.A 2: GoTo E
            Next
            mAyFmtStr = mAy1
        End If
    End With
Next
oAys = mAyFmtStr
Exit Function
R: ss.R
E: Fmt_Str_Repeat_ByAm_IntoAy = True: ss.B cSub, cMod, "pFmtStr,pAm", pFmtStr, jj.ToStr_Am(pAm)
End Function
Function Fmt_Str_Repeat_ByLm_IntoAy_Tst() As Boolean
Dim mFmtStr$
Dim mLm$
Dim mA$()
mFmtStr = "fkjsdf[{Itm}][{N}][{X}]lskdfj"
mLm = "Itm=Tbl;N=1,2,3;X=,x,xx,xxx"
If jj.Fmt_Str_Repeat_ByLm_IntoAy(mA, mFmtStr, mLm) Then Stop
Debug.Print Join(mA, vbCrLf & vbCrLf)
Shw_DbgWin
End Function
Function Fmt_Str_Repeat_ByLpAp_IntoAy(oAys$(), pFmtStr$, pLp$, ParamArray pAp()) As Boolean
'Aim: see Fmt_Str_Repeat_ByLpAp_IntoAy
Fmt_Str_Repeat_ByLpAp_IntoAy = Fmt_Str_Repeat_ByAm_IntoAy(oAys, pFmtStr, jj.Get_Am_ByLpVv(pLp, CVar(pAp)))
End Function
#If Tst Then
Function Fmt_Str_Repeat_ByLpAp_IntoAy_Tst() As Boolean
Dim mFmtStr$
mFmtStr = "INSERT INTO [$Ty{Itm}{N}{x}] ( NmTy{Itm}{N}{x} )" & _
" SELECT DISTINCT [#NmTy{Itm}{N}].NmTy{Itm}{N}{x}" & _
" FROM [#NmTy{Itm}{N}] LEFT JOIN [$Ty{Itm}{N}{x}] ON [#NmTy{Itm}{N}].NmTy{Itm}{N}{x} = [$Ty{Itm}{N}{x}].NmTy{Itm}{N}{x}" & _
" WHERE ((Not ([#NmTy{Itm}{N}].NmTy{Itm}{N}{x}) Is Null) AND (([$Ty{Itm}{N}{x}].NmTy{Itm}{N}{x}) Is Null));"
Dim mAyN$(): mAyN = Split(jj.Fmt_Str_Repeat(1, 2, "{N}", " "))
Dim mAyX(): mAyX = Array("X", "XX", "XXX")
Dim mAy$()
If jj.Fmt_Str_Repeat_ByLpAp_IntoAy(mAy, mFmtStr, "Itm,N,X", "Tbl", mAyN, mAyX) Then Stop
Debug.Print Join(mAy, vbCrLf & vbCrLf)
Shw_DbgWin
End Function
#End If
Function Fmt_Str_Repeat_ByLpAp$(pSepChr$, pFmtStr$, pLp$, ParamArray pAp())
Const cSub$ = "Fmt_Str_Repeat_ByLpAp"
Dim mAn$(): mAn = Split(pLp, ",")
Dim mN1%: mN1 = jj.Siz_Ay(mAn)
Dim mN2%: mN2 = jj.Siz_Vayv(CVar(pAp))
If mN1 <> mN2 Then ss.A 1, "pLp & pAp are not same size", , "pLp Siz,pAp Siz", mN1, mN2: GoTo E
Dim mS$: mS = pFmtStr
Dim J%
For J = 0 To mN1 - 1
    Dim mMacro$: mMacro = "{" & mAn(J) & "}"
    If VarType(pAp(J)) And vbArray Then
        Dim mAys$(): mAys = jj.Cv_AyV(pAp(J))
        mS = jj.Fmt_Str_Repeat_Ay(mS, mAys, , pSepChr, mMacro)
    Else
        mS = Replace(mS, mMacro, pAp(J))
    End If
Next
Fmt_Str_Repeat_ByLpAp = mS
Exit Function
R: ss.R
E: ss.B cSub, cMod, "pSepChr,pFmtStr,pLp,pAyPrm", pSepChr, pFmtStr, pLp, jj.ToStr_Vayv(CVar(pAp))
End Function
Function Fmt_Str_Repeat_ByLpAp_Tst() As Boolean
Dim mFmtStr$
mFmtStr = "INSERT INTO [$Ty{Itm}{N}{x}] ( NmTy{Itm}{N}{x} )" & _
" SELECT DISTINCT [#NmTy{Itm}{N}].NmTy{Itm}{N}{x}" & _
" FROM [#NmTy{Itm}{N}] LEFT JOIN [$Ty{Itm}{N}{x}] ON [#NmTy{Itm}{N}].NmTy{Itm}{N}{x} = [$Ty{Itm}{N}{x}].NmTy{Itm}{N}{x}" & _
" WHERE ((Not ([#NmTy{Itm}{N}].NmTy{Itm}{N}{x}) Is Null) AND (([$Ty{Itm}{N}{x}].NmTy{Itm}{N}{x}) Is Null));"
Dim mAyN$(): mAyN = Split(jj.Fmt_Str_Repeat(1, 2, "{N}", " "))
Dim mAyX(): mAyX = Array("X", "XX", "XXX")
Debug.Print jj.Fmt_Str_Repeat_ByLpAp(vbCrLf & vbCrLf, mFmtStr, "Itm,N,X", "Tbl", mAyN, mAyX)
Shw_DbgWin
End Function
Function Fmt_Ws(pRge As Range, pNRec&, Optional pRnoColrIdx& = 0) As Boolean
Const cSub$ = "Fmt_Ws"
Dim mStp!
On Error GoTo R
If pRge.Row >= 4 Then pRge(-2, 1).Value = "Export @ " & Now()

mStp = 1
'Set SubTotal
Dim mCnoLas As Byte: If jj.Fnd_CnoLas(mCnoLas, pRge(0, 1)) Then ss.A 1: GoTo E
Dim mWs As Worksheet: Set mWs = pRge.Parent
With pRge(1 + pNRec + 1, 2)
    If pNRec = 0 Then
        .Formula = 0
    Else
        Dim mCno As Byte: mCno = pRge.Column + 1
        Dim mRge As Range: Set mRge = mWs.Range(pRge(1, mCno), pRge(pNRec, mCno))
        Dim mAdr$: mAdr = mRge.Address(False, False)
        .Formula = jj.Q_S(mAdr, "=SUBTOTAL(3,*)")
    End If
    .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
End With

mStp = 2
'Copy Formula
If pNRec > 0 Then If jj.Cpy_Formula_ByCmt(pRge, pNRec) Then ss.A 2: GoTo E

mStp = 3
'MergeCells: If cell "A<cRnoDta> has color, it means it need to merge cells
If jj.Fmt_Ws_ByMge(pRge, mCnoLas, pNRec, pRge.Row - 1) Then ss.A 3: GoTo E

mStp = 4
'Use Row as color index
If pNRec > 0 Then If pRnoColrIdx > 0 Then If jj.Fmt_Ws_ByColr(pRge, pRnoColrIdx) Then ss.A 4: GoTo E

With pRge.Parent
    'If Not pNoExpTim Then .Range("A2").Value = "Exported @ " & Format(Now(), "yyyy/mm/dd hh:mm:ss")
    .Select
    .Activate
    .OutLine.ShowLevels , 1
End With
With pRge(1, 1)
    .Select
    .Activate
End With
Exit Function
R: ss.R
E: Fmt_Ws = True: ss.B cSub, cMod, "pWs,pNRec", jj.ToStr_Rge(pRge), pNRec
End Function
Function Fmt_Ws_ByColr(pRge As Range, pRnoColrIdx&) As Boolean
'Aim: Color some columns cell as indexed by {pRnoColrIdx}
'     Which column to Colr: Find mAyCno() & mAyColr() @ {pRnoColrIdx} if they have colour
'     Detect Last Column  : use also non-empty-cell @ cRnoNmFld
'     Detect Last Row     : non-empty-cell of column B
'     Starting Row        : gRnoDta
Const cSub$ = "Fmt_Ws_ByColr"
On Error GoTo R
If pRnoColrIdx <= 0 Then Exit Function
Dim mWs As Worksheet: Set mWs = pRge.Parent
jj.Shw_AllDta mWs
Dim iCno As Byte, mSq As New cSq
With mSq
    .Rno1 = pRge.Row
    .Rno2 = pRge.End(xlDown).Row
End With

Dim mAyCno() As Byte, mAyColr&()
If jj.Fnd_AyCnoColr(mAyCno, mAyColr, pRge, pRnoColrIdx) Then ss.A 1: GoTo E
Dim J%
For J = 0 To jj.Siz_Ay(mAyCno) - 1
    With mSq
        .Cno1 = mAyCno(J)
        .Cno2 = .Cno1
    End With
    Dim mRge As Range: If mSq.GetRge(mRge, mWs) Then ss.A 2: GoTo E
    mRge.Interior.Color = mAyColr(J)
Next
Exit Function
R: ss.R
E: Fmt_Ws_ByColr = True: ss.B cSub, cMod, "pRge,pRnoColrIdx", jj.ToStr_Rge(pRge), pRnoColrIdx
End Function
#If Tst Then
Function Fmt_Ws_ByColr_Tst() As Boolean
If jj.Cpy_Fil("P:\AppDef_Meta\MetaLgc.xls", "c:\Tmp\aa.xls", True) Then Stop: GoTo E
Dim mWb As Workbook: If jj.Opn_Wb_RW(mWb, "c:\tmp\aa.xls", , True) Then Stop: GoTo E
Dim mWs As Worksheet: Set mWs = mWb.Sheets("Stp")
If jj.Fmt_Ws_ByColr(mWs.Range("A5"), 3) Then Stop: GoTo E
Stop
GoTo X
E: Fmt_Ws_ByColr_Tst = True
X: jj.Cls_Wb mWb, False, True
End Function
#End If
Function Fmt_Ws_ByMge(pRge As Range, pCnoEnd As Byte, pNRec&, pRnoMgeIdx As Byte) As Boolean
'Aim: Mge some columns cell as indexed by {pRnoColrIdx}
'     Which column to Mge: Test the first cell of {pRnoMgeIdx} has color or not.
'                          If no color, no merge
'                          If Yes, use Col A as Key to merge cells.
'                          The columns to be merge is same color as the first cell. (mAyCno(mN%))
'     Detect Last Column : use also non-empty-cell @ cRnoNmFld
'     Detect Last Row    : non-empty-cell of column A
'     Starting Row       : gRnoDta
Const cSub$ = "Fmt_Ws_ByMege"
On Error GoTo R
Dim mWs As Worksheet: Set mWs = pRge.Parent
Dim mColr&: mColr = mWs.Cells(pRnoMgeIdx, pRge.Column).Interior.Color
If mColr = jj.g.cColrNo Then Exit Function
Dim iCno As Byte, mAyCno() As Byte, mN%
For iCno = pRge.Column To pCnoEnd
    Dim mRge As Range: Set mRge = mWs.Cells(pRnoMgeIdx, iCno)
    If mRge.Interior.Color = mColr Then
        ReDim Preserve mAyCno(mN): mAyCno(mN) = iCno: mN = mN + 1
    End If
Next
Dim iRno&, mRnoBeg&
mRnoBeg = pRge.Row
Dim mCnoLookAt As Byte: mCnoLookAt = pRge.Column
Dim mV, mVLas: mVLas = mWs.Cells(mRnoBeg, mCnoLookAt)
For iRno = mRnoBeg To mRnoBeg + pNRec - 1
    If iRno Mod 50 = 0 Then jj.Shw_Sts "Merging cell at row " & iRno & "..."
    mV = mWs.Cells(iRno, mCnoLookAt).Value
    If mV <> mVLas Then
        If jj.Fmt_Ws_MgeCells(mWs, mRnoBeg, iRno - 1, mAyCno) Then ss.A 2: GoTo E
        mRnoBeg = iRno
        mVLas = mV
    End If
Next
GoTo X
R: ss.R
E: Fmt_Ws_ByMge = True: ss.B cSub, cMod, "pRge,pRnoMgeIdx", jj.ToStr_Rge(pRge), pRnoMgeIdx
X: jj.Clr_Sts
End Function
#If Tst Then
Function Fmt_Ws_ByMge_Tst() As Boolean
If jj.Cpy_Fil("P:\AppDef_Meta\MetaLgc.xls", "c:\Tmp\aa.xls", True) Then Stop: GoTo E
Dim mWb As Workbook: If jj.Opn_Wb_RW(mWb, "c:\tmp\aa.xls", , True) Then Stop: GoTo E
Dim mWs As Worksheet: Set mWs = mWb.Sheets("OldQsT")
Dim mRge As Range: Set mRge = mWs.Range("C4")
Dim mCnoLas As Byte: If jj.Fnd_CnoLas(mCnoLas, mRge) Then Stop: GoTo E
Dim mRnoLas&: If jj.Fnd_RnoLas(mRnoLas, mRge) Then Stop: GoTo E
Dim mNRec&: mNRec = mRnoLas - 5
If jj.Fmt_Ws_ByMge(mRge, mCnoLas, mNRec, 4) Then Stop: GoTo E
Exit Function
E: Fmt_Ws_ByMge_Tst = True
End Function
#End If
Function Fmt_Ws_MgeCells(pWs As Worksheet, pRnoBeg&, pRnoEnd&, pAyCno() As Byte) As Boolean
'Aim: for each {pCno} merge cells vertically from {pRnoBeg} to {pRnoEnd}
Const cSub$ = "Fmt_Ws_MgeCells"
If pRnoBeg >= pRnoEnd Then Exit Function

On Error GoTo R
Dim iCno As Byte, J%, mRge As Range
Dim mXls As Excel.Application: Set mXls = pWs.Application: mXls.ScreenUpdating = False: mXls.DisplayAlerts = False
On Error GoTo R
For J = 0 To jj.Siz_Ay(pAyCno) - 1
    Set mRge = pWs.Range(pWs.Cells(pRnoBeg, pAyCno(J)), pWs.Cells(pRnoEnd, pAyCno(J)))
    With mRge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .MergeCells = False
        .Merge
    End With
Next
GoTo X
R: ss.R
E: Fmt_Ws_MgeCells = True: ss.B cSub, cMod, "pWs,pRnoBeg,pRnoEnd,pAyCno", jj.ToStr_Ws(pWs), pRnoBeg, pRnoEnd, jj.ToStr_AyByt(pAyCno)
X: mXls.ScreenUpdating = True
   mXls.DisplayAlerts = True
End Function
Function Fmt_Str_Repeat_Lv$(pFmtStr$, pLv$, Optional pSepChr$ = cComma, Optional pMacroStr$ = "{B}")
'Aim: return a string joining the string after each substitute {pMacroStr} in {pFmtStr} by the list of values in {pLv}
Const cSub$ = "Fmt_Str_Repeat_Lv"
Dim mAn$(): mAn = Split(pLv, cComma)
Fmt_Str_Repeat_Lv = jj.Fmt_Str_Repeat_Ay(pFmtStr, mAn, , pSepChr, pMacroStr)
End Function
#If Tst Then
Function Fmt_Str_Repeat_Lv_Tst() As Boolean
Debug.Print Fmt_Str_Repeat_Lv("Tbl{B}", "x,xx,xxx")
Debug.Print Fmt_Str_Repeat_Lv("Tbl{B}", "x,xx,xxx", vbLf)
End Function
#End If
Function Fmt_Str_Repeat_Ay_MultiLine$(pMultiLines$, pAys$(), Optional pN% = 0, Optional pMacroStr$ = "{N}")
'Aim: Return a string of multiple lines.  The line contains {N} will be repeated by substitue {N} by pAn$(..)
Const cSub$ = "Fmt_Str_Repeat_Ay_MultiLine"
Dim mAyL$(): mAyL = Split(pMultiLines, vbCrLf)
Dim mL$, iL%
For iL = 0 To jj.Siz_Ay(mAyL) - 1
    mL = jj.Add_Str(mL, Fmt_Str_Repeat_Ay(mAyL(iL), pAys, pN, vbCrLf, pMacroStr), vbCrLf)
Next
Fmt_Str_Repeat_Ay_MultiLine = mL
End Function
#If Tst Then
Function Fmt_Str_Repeat_Ay_MultiLine_Tst() As Boolean
Dim mLines$:
mLines = "Line1:lksdjflskdf sdklfj" & vbCrLf & _
"Line2:klsdjf{B}klsdjf" & vbCrLf & _
"Line3:ksldjfslkdf"
Dim mAys$(): mAys = Split("<Itm1>,<Itm2>,<Itm3>,<Itm4>", ",")
Debug.Print Fmt_Str_Repeat_Ay_MultiLine(mLines, mAys, , "{B}")
End Function
#End If
Function Fmt_Str_Repeat_Ay_IntoAy(oAys$(), pFmtStr$, pAys$(), Optional pN% = 0, Optional pMacroStr$ = "{N}") As Boolean
'Aim: Format a string with {N} by repeatly join it after substitue {N} by pAn$(0 to pN)
Const cSub$ = "Fmt_Str_Repeat_Ay"
If InStr(pFmtStr, pMacroStr) = 0 Then
    ReDim oAys(0): oAys(0) = pFmtStr
    Exit Function
End If
Dim mN%
If pN = 0 Then
    mN = jj.Siz_Ay(pAys)
Else
    mN = pN
End If
Dim J%, mA$
ReDim oAys(mN - 1)
For J = 0 To mN - 1
    oAys(J) = Replace(pFmtStr, pMacroStr, pAys(J))
Next
Exit Function
E: Fmt_Str_Repeat_Ay_IntoAy = True: ss.B cSub, cMod, "pFmtStr,pAys,pN,pMacrosStr", pFmtStr, Join(pAys, ","), pN, pMacroStr
End Function
Function Fmt_Str_Repeat_Ay$(pFmtStr$, pAys$(), Optional pN% = 0, Optional pSepChr$ = cComma, Optional pMacroStr$ = "{N}")
'Aim: Format a string with {N} by repeatly join it after substitue {N} by pAn$(0 to pN)
Const cSub$ = "Fmt_Str_Repeat_Ay"
If InStr(pFmtStr, "{") = 0 Then Fmt_Str_Repeat_Ay = pFmtStr: Exit Function
Dim mN%
If pN = 0 Then
    mN = jj.Siz_Ay(pAys)
Else
    mN = pN
End If
Dim J%, mA$
For J = 0 To mN - 1
    mA = jj.Add_Str(mA, Replace(pFmtStr, pMacroStr, pAys(J)), pSepChr)
Next
Fmt_Str_Repeat_Ay = mA
End Function
#If Tst Then
Function Fmt_Str_Repeat_Ay_Tst() As Boolean
Dim mA$(3)
Dim J%
For J = 0 To 3
    mA(J) = "[" & J & "]"
Next
Debug.Print jj.Fmt_Str_Repeat_Ay("xxxx{N}yyyyy", mA, 2, " and ")
End Function
#End If
Function Fmt_Sql_InFbTar$(pSql$, pFbTar$)
'Aim:   Fmt
'           select * into [{xxx@FbTar}] from ..
'       Into
'           Select * into [xxx] in 'nnnnn' from ..      If pFbTar is nnnnn
'       or
'           Select * into [xxx] from ..                 If pFbTar is empty
Dim p2%: p2% = InStr(pSql, "@FbTar}]")
'select * into [{xxx@FbTar}] from ..
'              ^P1  ^P2     ^P2+8

'MID(xx,P1+2,P2-P1-2)
'=               xxx

'Left(pSql, P1 - 1)
'MID(pSql,P2-3)
Dim mR$: mR = pSql
If p2 <= 0 Then GoTo X
Dim p1%: p1 = InStrRev(Left(pSql, p2 - 1), "[{")
If p1 <= 0 Then GoTo X
Dim mInFb$: mInFb = jj.Cv_Fb2InFb(pFbTar)
Dim mX$: mX$ = "[" & mID(pSql, p1 + 2, p2 - p1 - 2) & "]" & mInFb
mR = Left(pSql, p1 - 1) & mX & mID(pSql, p2 + 8)
X: Fmt_Sql_InFbTar = mR
End Function
#If Tst Then
Function Fmt_Sql_InFbTar_Tst() As Boolean
Const cSub$ = "Fmt_Sql_InFbTar_Tst"
Dim mSql$, mFbTar$
mFbTar = "c:\aa.mdb"
mSql = "select * into [{xxx@FbTar}] from .."
Dim mRetSql$: mRetSql = jj.Fmt_Sql_InFbTar(mSql, mFbTar)
jj.Shw_Dbg cSub, cMod, "mSql,mFbTar,RetSql", mSql, mFbTar, mRetSql
End Function
#End If
Function Fmt_Border(pRge As Range, Optional pLinSty As XlLineStyle = XlLineStyle.xlContinuous, Optional pBdrWgt As XlBorderWeight = XlBorderWeight.xlMedium) As Boolean
Const cSub$ = "Fmt_Border"
pRge.BorderAround pLinSty, pBdrWgt
Dim mSq As tSq: If jj.Cv_Rge2Sq(mSq, pRge) Then ss.A 1: GoTo E
Dim mRge As Range

Set mRge = jj.Crt_Rge_HLin(pRge.Worksheet, mSq.r2 + 1, mSq.c1, mSq.c2)
With mRge.Borders(xlEdgeTop)
    .LineStyle = pLinSty
    .Weight = pBdrWgt
End With

Set mRge = jj.Crt_Rge_VLin(pRge.Worksheet, mSq.c2 + 1, mSq.r1, mSq.r2)
With mRge.Borders(xlEdgeLeft)
    .LineStyle = pLinSty
    .Weight = pBdrWgt
End With
Exit Function
R: ss.R
E: Fmt_Border = True: ss.B cSub, cMod, "pRge,pLinSty,pBdrWgt", jj.ToStr_Rge(pRge), pLinSty, pBdrWgt
End Function
Function Fmt_Str$(pFmtStr$, ParamArray pAp())
Dim mS$, mA$, mP%, I As Byte, iPrm
mS = Replace(pFmtStr, "|", vbCrLf)
I = 0
For Each iPrm In pAp
    mS = Replace(mS, "{" & I & "}", Nz(iPrm, "Null")): I = I + 1
Next
Fmt_Str = mS
End Function
#If Tst Then
Function Fmt_Str_Tst() As Boolean
Debug.Print jj.Fmt_Str("dsf{0},{1}dlf", 1, Now)
End Function
#End If
Function Fmt_Str_Repeat$(pBeg As Byte, pN As Byte, Optional pFmtStr$ = "{N}", Optional pSepChr$ = cComma, Optional pMacroStr$ = "{N}")
'Aim: Build a string to repeating {pFmtStr} {pN} times from {pBeg} with separated by {pSepChr}.  {pFmtStr} has {N} as the Idx.
Const cSub$ = "Fmt_Str_Repeat"
Dim mA$, J As Byte
For J = pBeg To pBeg + pN - 1
    mA = jj.Add_Str(mA, Replace(pFmtStr, pMacroStr, J), pSepChr)
Next
Fmt_Str_Repeat = mA
End Function
#If Tst Then
Function Fmt_Str_Repeat_Tst() As Boolean
Dim mExpr$
mExpr = "Fmt_Str_Repeat(0, 10, ""a{N} as xx{N}"")"
Debug.Print "================="
Debug.Print mExpr
Debug.Print Eval(mExpr) ' jj.Fmt_Str_Repeat(1, 10, "a{N} as xx{N}")
Debug.Print
mExpr = "Fmt_Str_Repeat(1, 10, ""a{N} as xx{N}"")"
Debug.Print mExpr
Debug.Print Eval(mExpr) ' jj.Fmt_Str_Repeat(1, 10, "a{N} as xx{N}")
End Function
#End If
Function Fmt_Str_ByAyKV$(pFmtStr$, pAyK$(), pAyV$(), Optional pExclSqBkt As Boolean = False)
Const cSub$ = "Fmt_Str_ByAyKV"
Dim mR$: mR = pFmtStr
If InStr(pFmtStr, "{") = 0 Then GoTo X
Dim N1%: N1 = jj.Siz_Ay(pAyK)
Dim N2%: N2 = jj.Siz_Ay(pAyV)
If N1 <> N2 Then ss.A 1, "Count in pLn & pAp are different", , "Cnt in pAyK, Cnt in pAyV", N1, N2: GoTo E
Dim J%: For J = 0 To N1 - 1
    If InStr(mR, "{") <= 0 Then GoTo X
    If Not pExclSqBkt Then mR = Replace(mR, "[{" & pAyK(J) & "}]", pAyV(J))
    mR = Replace(mR, "{" & pAyK(J) & "}", pAyV(J))
Next
GoTo X
E: ss.B cSub, cMod, "pFmtStr,pAyK,pAyV,pExclSqlBkt", pFmtStr, jj.ToStr_Ays(pAyK), jj.ToStr_Ays(pAyV), pExclSqBkt
X: Fmt_Str_ByAyKV = mR
End Function
#If Tst Then
Function Fmt_Str_ByAyKV_Tst() As Boolean
Dim mAyK$(1), mAyV$(2)
Debug.Print Fmt_Str_ByAyKV("lksd{a}jf", mAyK, mAyV)
End Function
#End If
Function Fmt_Str_ByAm$(pFmtStr$, pAm() As tMap, Optional pExclSqBkt As Boolean = False)
Dim N%: N = jj.Siz_Am(pAm)
Dim mR$: mR = pFmtStr
If N = 0 Then GoTo X
ReDim mAyK$(N - 1), mAyV$(N - 1)
Dim J%
For J = 0 To N - 1
    With pAm(J)
        mAyK(J) = .F1
        mAyV(J) = .F2
    End With
Next
mR = jj.Fmt_Str_ByAyKV(pFmtStr, mAyK, mAyV, pExclSqBkt)
X:
    Fmt_Str_ByAm = mR
End Function
Function Fmt_Str_ByColl$(pFmtStr$, pMacroList As VBA.Collection)
Const cSub$ = "Fmt_Str_ByColl"
Dim mR$: mR = pFmtStr
If jj.IsNothing(pMacroList) Then GoTo X
Dim AyK$(), AyV$()
If jj.Cv_CollKvStr_To2Ay(pMacroList, AyK, AyV) Then ss.A 1: GoTo E
mR = jj.Fmt_Str_ByAyKV(pFmtStr, AyK, AyV)
GoTo X
E: ss.B cSub, cMod, "pFmtStr,pMacroList", pFmtStr, jj.ToStr_Coll(pMacroList)
X:
    Fmt_Str_ByColl = mR
End Function
Function Fmt_Str_ByLpAp_ExclSqBkt$(pFmtStr$, pLp$, ParamArray pAp())
'Aim: use {xx} as the substitue string to substitue {pFmtStr} by the named list in {pLp} & {pAp}
Const cSub$ = "Fmt_Str_ByLpAp"
Dim mAm() As tMap: If jj.Brk_LpVv2Am(mAm, pLp, CVar(pAp)) Then ss.A 1: GoTo E
Fmt_Str_ByLpAp_ExclSqBkt = Fmt_Str_ByAm(pFmtStr, mAm, True)
Exit Function
E: ss.B cSub, cMod, "pFmtStr,pLp,pAp", pFmtStr, pLp, jj.ToStr_Vayv(CVar(pAp))
End Function
Function Fmt_Str_ByLpAp$(pFmtStr$, pLp$, ParamArray pAp())
'Aim: use [{xx}] or {xx} as the substitue string to substitue {pFmtStr} by the named list in {pLp} & {pAp}
Const cSub$ = "Fmt_Str_ByLpAp"
Dim mAm() As tMap: If jj.Brk_LpVv2Am(mAm, pLp, CVar(pAp)) Then ss.A 1: GoTo E
Fmt_Str_ByLpAp = Fmt_Str_ByAm(pFmtStr, mAm)
Exit Function
E: ss.B cSub, cMod, "pFmtStr,pLp,pAp", pFmtStr, pLp, jj.ToStr_Vayv(CVar(pAp))
End Function
#If Tst Then
Function Fmt_Str_ByLpAp_Tst() As Boolean
MsgBox jj.Fmt_Str_ByLpAp("abc{Nm}_x_[{Date}]_y_{String}", "Nm,Date,String", 123, Date, "klsjdf")
End Function
#End If
Function Fmt_Str_ByLm$(pFmtStr$, pLm$, Optional pExclSqBkt As Boolean = False)
Const cSub$ = "Fmt_Str_ByLm"
Dim mAm() As tMap: mAm = jj.Get_Am_ByLm(pLm)
Fmt_Str_ByLm = Fmt_Str_ByAm(pFmtStr, mAm, pExclSqBkt)
Exit Function
E: ss.B cSub, cMod, "pFmtStr,pLm,pExclSqBkt", pFmtStr, pLm, pExclSqBkt
End Function
#If Tst Then
Function Fmt_Str_ByLm_Tst() As Boolean
Dim mLm$: mLm = "abc=---abc---,def=---def---"
Dim mS$: mS = "adf {abc} xxx {def}"
Debug.Print mS
Debug.Print jj.Fmt_Str_ByLm(mS, mLm)
Shw_DbgWin
End Function
#End If
Function Fmt_Str_ByRs$(pFmtStr$, pRs As DAO.Recordset)
Const cSub$ = "Fmt_Str_ByRs"
'Aim: pFmtStr is in format of xxxx{Fld1}xxx{Fld2}.  Return the subst string by subst the fields in {pRs} into {pFmtStr}
Dim mS$: mS = pFmtStr
Dim J%: For J = 0 To pRs.Fields.Count - 1
    Dim mA$: mA = "{" & pRs.Fields(J).Name & "}"
    Dim mP%: mP = InStr(mS, mA)
    Do While mP > 0
        If IsNull(pRs.Fields(J).Value) Then ss.A 1, "The field value of given Rs is null", , "Rs,Field Nam with Null value", jj.ToStr_Rs(pRs), pRs.Fields(J).Name: Exit Do
        mS = Replace(mS, mA, pRs.Fields(J).Value)
        mP = InStr(mS, mA)
    Loop
Next
Fmt_Str_ByRs = mS
End Function
Function Fmt_Tbl(pQt As QueryTable) As Boolean
Const cSub$ = "Fmt_Tbl"
'0 Read Definition
Dim mDef As tFmtTbl_Def: If jj.Read_Def_FmtTbl(mDef, pQt) Then ss.A 1: GoTo E

Dim mCno As Byte
'1 Fmt VLine
''Fmt VLineLeft
Dim mWs As Worksheet: Set mWs = pQt.Parent
Dim iCno%: For iCno = 0 To jj.Siz_Ay(mDef.AyCnoVLine_Left) - 1
    mCno = mDef.AyCnoVLine_Left(iCno)
    jj.Crt_Rge_VLin(mWs, mCno, mDef.SqDta.r1, mDef.SqDta.r2).Borders(xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
Next
''Fmt VLineRight
For iCno = 0 To jj.Siz_Ay(mDef.AyCnoVLine_Right) - 1
    mCno = mDef.AyCnoVLine_Right(iCno)
    jj.Crt_Rge_VLin(mWs, mCno, mDef.SqDta.r1, mDef.SqDta.r2).Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
Next
''Fmt VLineLeftMedium
For iCno = 0 To jj.Siz_Ay(mDef.AyCnoVLine_LeftMedium) - 1
    mCno = mDef.AyCnoVLine_LeftMedium(iCno)
    With jj.Crt_Rge_VLin(mWs, mCno, mDef.SqDta.r1, mDef.SqDta.r2).Borders(xlEdgeLeft)
        .LineStyle = XlLineStyle.xlContinuous
        .Weight = XlBorderWeight.xlMedium
    End With
Next
''Fmt VLineRightMedium
For iCno = 0 To jj.Siz_Ay(mDef.AyCnoVLine_RightMedium) - 1
    mCno = mDef.AyCnoVLine_RightMedium(iCno)
    With jj.Crt_Rge_VLin(mWs, mCno, mDef.SqDta.r1, mDef.SqDta.r2).Borders(xlEdgeRight)
        .LineStyle = XlLineStyle.xlContinuous
        .Weight = XlBorderWeight.xlMedium
    End With
Next

'2 Merge
Dim mSq As tSq: mSq = mDef.SqDta
Dim mRge As Range
If pQt.FieldNames Then mSq.r1 = mSq.r1 + 1
Dim N%: N% = jj.Siz_AyRgeCno(mDef.AyRgeCno_Merge)
If N% > 0 Then
    gXls.DisplayAlerts = False
    mSq = mDef.SqDta
    Dim J%: For J = 0 To N - 1
        Dim iRgeCno As tRgeCno: iRgeCno = mDef.AyRgeCno_Merge(J)
        With mSq
            .c1 = iRgeCno.Fm
            .c2 = iRgeCno.To
            .r1 = mDef.SqDta.r1 + IIf(mDef.Qt.FieldNames, 1, 0)
            .r2 = mDef.SqDta.r2
        End With
        jj.Crt_Rge_FmSq(mWs, mSq).Merge True
    Next
    gXls.DisplayAlerts = True
End If

'3 [SepLin] Set Separating Lines
Dim iRno&
Dim mRnoBeg&: mRnoBeg = mDef.SqDta.r1 + IIf(pQt.FieldNames, 1, 0)
Dim mRnoEnd&: mRnoEnd = mDef.SqDta.r2
If mDef.IsSepLin Then
    ''Draw first line @ iRno = r1 + 3
    mSq = mDef.SqDta
    iRno = mRnoBeg + 3
    If mSq.r2 > iRno Then jj.Crt_Rge_HLin(mWs, iRno, mSq.c1, mSq.c2).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlDot
    ''Cpy first 3 line format downwards
    If mSq.r2 - 1 >= iRno Then
        Dim mSq1 As tSq: mSq1 = mSq: mSq1.r1 = iRno: mSq1.r2 = iRno + 2
        Dim mSq2 As tSq: mSq2 = mSq: mSq2.r1 = iRno + 3: mSq2.r2 = mSq.r2 - 1
        jj.Crt_Rge_FmSq(mWs, mSq1).Copy
        jj.Crt_Rge_FmSq(mWs, mSq2).PasteSpecial xlPasteFormats
    End If
End If

'4.1 SubTot
For iCno = 0 To jj.Siz_Ay(mDef.AyCnoSubTot) - 1
    mCno = mDef.AyCnoSubTot(iCno)
    Set mRge = jj.Crt_Rge_VLin(mWs, mCno, mRnoBeg, mRnoEnd)
    mWs.Cells(mSq.r2 + 1, mCno).Formula = jj.Fmt_Str("=Subtotal(9,{0})", mRge.Address(, False))
Next

'4.1 Avg
For iCno = 0 To jj.Siz_Ay(mDef.AyCnoAvg) - 1
    mCno = mDef.AyCnoAvg(iCno)
    Set mRge = jj.Crt_Rge_VLin(mWs, mCno, mRnoBeg, mRnoEnd)
    mWs.Cells(mSq.r2 + 1, mCno).Formula = jj.Fmt_Str("=SubTotal(1,{0})", mRge.Address(, False))
Next

'4.2 Count
mCno = mDef.CnoCnt
If mCno > 0 Then
    Set mRge = jj.Crt_Rge_VLin(mWs, mCno, mRnoBeg, mRnoEnd)
    With mWs.Cells(mSq.r2 + 1, mCno)
        .Formula = jj.Fmt_Str("=Subtotal(3,{0})", mRge.Address(, False))
        .NumberFormat = "#,##0"
    End With
End If

'4.3 Multi-Lvl Border & Color
If mDef.NLvl >= 2 Then If jj.Fmt_Tbl_NLvl_Border_n_Color(mWs, mDef) Then ss.A 1: GoTo E

'5 Copy formula
For iCno = 0 To jj.Siz_Ay(mDef.AyCnoFormula) - 1
    mCno = mDef.AyCnoFormula(iCno)
    With mWs.Cells(mRnoBeg, mCno)
        .Formula = mDef.AyFormula(iCno)
        .Copy
    End With
    Set mRge = jj.Crt_Rge_VLin(mWs, mCno, mRnoBeg, mRnoEnd)
    With mRge
        .PasteSpecial xlPasteFormulas
        .NumberFormat = mWs.Cells(mDef.SqDta.r1, mCno).NumberFormat
    End With
Next

'6 Border Around
jj.Fmt_Border jj.Crt_Rge_FmSq(mWs, mDef.SqDta)

'7 Remove Fmt Definition
mWs.Range(mDef.Sq.r1 & ":" & mDef.Sq.r2).Delete
mWs.Columns(mDef.Sq.c2).ColumnWidth = 1
jj.Crt_Rge_Col(mWs, mDef.Sq.c2, mDef.Sq.c2).ClearContents
jj.Crt_Rge_Col(mWs, mDef.Sq.c2 + 1, 255).Delete

mWs.Activate
mWs.Cells(1, 1).Select
mWs.OutLine.ShowLevels 1, 1
GoTo X
R: ss.R
E: Fmt_Tbl = True: ss.B cSub, cMod, "pQt", jj.ToStr_Qt(pQt)
X: jj.Clr_Sts
End Function
Function Fmt_Tbl_NLvl_Border_n_Color(pWs As Worksheet, pDef As tFmtTbl_Def) As Boolean
Const cSub$ = "Fmt_Tbl_NLvl_Border_n_Color"
If pDef.NLvl < 2 Then Exit Function

Dim mFb_FmtLvl$, mNmt_FmtLvl$
If jj.Fnd_Fb_FmCnnStr(mFb_FmtLvl, pDef.Qt.Connection) Then ss.A 1: GoTo E
If jj.Fnd_Nmt_FmQt(mNmt_FmtLvl, pDef.Qt) Then ss.A 2: GoTo E
mNmt_FmtLvl = mNmt_FmtLvl & "_FmtLvl"

'Find mDb, NLvl, mAyRgeCnoLvl, mAyColr()
Dim J As Byte, c1 As Byte, c2 As Byte, r1&, r2&
Dim mDb As DAO.Database: If jj.Opn_Db_R(mDb, mFb_FmtLvl) Then ss.A 3: GoTo E
Dim NLvl As Byte: If jj.Fnd_MaxVal(NLvl, mNmt_FmtLvl, "Lvl", , mDb) Then ss.A 4: GoTo E
If NLvl <= 0 Then ss.A 5, "Max(Lvl) in mNmt_FmtLvl cannot be 0", eRunTimErr, "Max(Lvl),mNmt_FmtLvl,mDb", NLvl, mNmt_FmtLvl, jj.ToStr_Db(mDb): GoTo E
If NLvl > pDef.NLvl Then ss.A 6, "NLvl found in {Nmt_FmtLvl} & NLvl > pDef", eRunTimErr, "NLvl in [Nmt_FmtLvl],NLvl in Def,Fb_FmtLvl,Nmt_FmtLvl", NLvl, pDef.NLvl, mFb_FmtLvl, mNmt_FmtLvl: GoTo E
ReDim mAyRgeCnoLvl(0 To NLvl - 1) As tRgeCno
ReDim mAyColr&(0 To NLvl - 1)

With pDef
    If NLvl >= 1 Then mAyRgeCnoLvl(0) = .RgeCno_Lvl1
    If NLvl >= 2 Then mAyRgeCnoLvl(1) = .RgeCno_Lvl2
    If NLvl >= 3 Then mAyRgeCnoLvl(2) = .RgeCno_Lvl3
    If NLvl >= 4 Then mAyRgeCnoLvl(3) = .RgeCno_Lvl4
    'Find mRowOffSet
    Dim mRowOffSet As Byte: mRowOffSet = .SqDta.r1 + IIf(.Qt.FieldNames, 0, -1)
End With
r1 = pDef.SqDta.r1 - 1
If NLvl >= 1 Then c1 = mAyRgeCnoLvl(0).Fm: mAyColr(0) = pWs.Cells(r1, c1).Interior.ColorIndex
If NLvl >= 2 Then c1 = mAyRgeCnoLvl(1).Fm: mAyColr(1) = pWs.Cells(r1, c1).Interior.ColorIndex
If NLvl >= 3 Then c1 = mAyRgeCnoLvl(2).Fm: mAyColr(2) = pWs.Cells(r1, c1).Interior.ColorIndex
If NLvl >= 4 Then c1 = mAyRgeCnoLvl(3).Fm: mAyColr(3) = pWs.Cells(r1, c1).Interior.ColorIndex

Dim RnoLas&: RnoLas = pDef.SqDta.r2
For J = 0 To NLvl - 1
    Dim mSql$: mSql = jj.Fmt_Str("Select Rno_Beg, Rno_End from {0} where Lvl={1} Order by Rno_Beg", mNmt_FmtLvl, J + 1)
    With mDb.OpenRecordset(mSql)
        Dim L&
        While Not .EOF
            L = L + 1
            c1 = mAyRgeCnoLvl(J).Fm: c2 = pDef.SqDta.c2
            r1 = !Rno_Beg + mRowOffSet
            r2 = !Rno_End + mRowOffSet
            If L Mod 50 = 0 Then VBA.DoEvents: SysCmd acSysCmdSetStatus, "Lvl" & J & "(" & NLvl - 1 & ")  Line " & r1 & " of " & RnoLas&
            
            If J < NLvl - 1 Then If r1 <> r2 Then pWs.Rows(r1 + 1 & ":" & r2).Group

            'If jj.Clr_InnerBorder(jj.Crt_Rge_FmSq(pWs, c1, r1, c2, r2)) Then jj.Fmt_Tbl_NLvl_Border_n_Color: ss.xx 1, cSub, cMod: GoTo E
            jj.Crt_Rge_Fm2Pts(pWs, r1, c1, r2, c2).Interior.ColorIndex = xlNone
            c2 = mAyRgeCnoLvl(J).To
            With jj.Crt_Rge_Fm2Pts(pWs, r1, c1, r2, c2)
                .BorderAround XlLineStyle.xlContinuous, xlThin
                .Interior.ColorIndex = mAyColr(J)
            End With
            .MoveNext
        Wend
        .Close
    End With
Next
GoTo X
R: ss.R
E: Fmt_Tbl_NLvl_Border_n_Color = True: ss.B cSub, cMod
X: jj.Clr_Sts
   jj.Cls_Db mDb
End Function
Function Fmt_Tbl_Tst() As Boolean
Const cSub$ = "Fmt_Tbl_Tst"
Dim mFfnTo$, mFfnFm$, mWb As Workbook, iWs As Worksheet, iQt As QueryTable
Dim mCase%: mCase = 3
Select Case mCase
Case 1
    mFfnTo = jj.Sdir_Tmp & "a.xls"
    mFfnFm = jj.Sffn_Tp("ARInq")
    If jj.Cpy_Fil(mFfnFm, mFfnTo, True) Then ss.A 1: GoTo E
    If jj.Opn_Wb_RW(mWb, mFfnTo) Then ss.A 2: GoTo E
    For Each iWs In mWb.Sheets
        For Each iQt In iWs.QueryTables
            jj.Shw_Dbg cSub, cMod, "Test for Formatting a query table in Xls", "Result of formatting Ws, iWs, iQt", jj.Fmt_Tbl(iQt), iWs.Name, iQt.Name
        Next
    Next
Case 2
    mFfnTo = jj.Sdir_Tmp & "a.xls"
    mFfnFm = jj.Sffn_Tp("CusLstForEdt")
    If jj.Cpy_Fil(mFfnFm, mFfnTo, True) Then ss.A 3: GoTo E
    If jj.Opn_Wb_RW(mWb, mFfnTo) Then ss.A 4: GoTo E
    For Each iWs In mWb.Sheets
        For Each iQt In iWs.QueryTables
            jj.Shw_Dbg cSub, cMod, "Test for Formatting a query table in Xls", "Result of formatting Ws, iWs, iQt", jj.Fmt_Tbl(iQt), iWs.Name, iQt.Name
        Next
    Next
Case 3
    If jj.Opn_Wb_RW(mWb, "c:\tmp\aa.xls") Then ss.A 4: GoTo E
    For Each iWs In mWb.Sheets
        For Each iQt In iWs.QueryTables
            jj.Shw_Dbg cSub, cMod, "Test for Formatting a query table in Xls", "Result of formatting Ws, iWs, iQt", jj.Fmt_Tbl(iQt), iWs.Name, iQt.Name
        Next
    Next
End Select
gXls.Visible = True
Stop
GoTo X
R: ss.R
E: Fmt_Tbl_Tst = True: ss.B cSub, cMod
X: jj.Cls_Wb mWb, False, True
End Function
Function Fmt_Wb(pWb As Workbook) As Boolean
Const cSub$ = "Fmt_Wb"
Dim mMsg$
Dim iWs As Worksheet: For Each iWs In pWb.Worksheets
    Dim iQt As QueryTable: For Each iQt In iWs.QueryTables
        If Fmt_Tbl(iQt) Then mMsg = jj.Add_Str(mMsg, iWs.Name)
    Next
Next
If mMsg <> "" Then ss.A 1, "Following Ws has error during formatting", "Ws List", mMsg: GoTo E
Exit Function
E:
: ss.B cSub, cMod, "pWb", jj.ToStr_Wb(pWb)
    Fmt_Wb = True
End Function
Function Fmt_Wb_Tst() As Boolean
Const cSub$ = "Fmt_Wb_Tst"
Dim mFfn$:
'mFfn = "N:\MONTHEND\PriceOverride TradeOffer SalesDiscount\Reports\PriceOverride_Yr_2006.xls"
mFfn = "N:\MONTHEND\PriceOverride TradeOffer SalesDiscount\Reports\SalesDiscount_2006"
Dim mWb As Workbook: If jj.Opn_Wb_RW(mWb, mFfn) Then ss.A 1: GoTo E
If jj.Fmt_Wb(mWb) Then ss.A 2: GoTo E
gXls.Visible = True
Exit Function
R: ss.R
E: Fmt_Wb_Tst = True: ss.B cSub, cMod
End Function
Function Fmt_WsOL_ByCol(pRge As Range, Optional pCithOL As Byte = 1, Optional pCithIns As Byte = 0) As Boolean
'Aim: Set outline level of data at {pRge} by the cells content.  Assume the cell content is outline level
'     Optionally insert cell and shift right at pCnoIns relative to pRge
Const cSub$ = "Fmt_WsOL_ByCol"
On Error GoTo R
Dim mWs As Worksheet: Set mWs = pRge.Parent
Dim iRno&, mLvl As Byte, mV

Dim mRnoLas&: If jj.Fnd_RnoLas(mRnoLas, pRge) Then ss.A 1: GoTo E
If pCithIns > 0 Then
    Dim mCnoLas As Byte: mCnoLas = pRge(0, 1).End(xlToRight).Column
    Dim mRgeBlk As Range: Set mRgeBlk = pRge.Range(pRge(1, 1), mWs.Cells(mRnoLas, mCnoLas))
    mRgeBlk.Sort Key1:=pRge(1, pCithOL), Order1:=xlAscending, Header:=xlNo _
        , MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Dim mAyRgeRno() As tRgeRno: If jj.Fnd_AyRgeRno(mAyRgeRno, pRge(1, pCithOL)) Then ss.A 3: GoTo E
    Dim J%
    Dim mCnoOL As Byte: mCnoOL = pRge.Column + pCithOL - 1
    Dim mCnoIns As Byte: mCnoIns = pRge.Column + pCithIns - 1
    For J = 0 To jj.Siz_AyRgeRno(mAyRgeRno) - 1
        With mAyRgeRno(J)
            mV = mWs.Cells(.Fm, mCnoOL).Value
            If VarType(mV) <> vbDouble Then ss.A 4, "The data type of mAyRgeRno(J).Fm content is not Double, which is used as OutLine", , "mAyRgeRno(J).Fm", .Fm: GoTo E
            If 0 > mV Or mV > 15 Then ss.A 5, "The value of mAyRgeRno(J).Fm content is not between 0 to 15", , "mAyRgeRno(J).Fm,The Val", .Fm, mV: GoTo E
            mLvl = mV
            If mLvl >= 1 Then
                Dim mRge As Range: Set mRge = mWs.Range(mWs.Cells(.Fm, mCnoIns), mWs.Cells(.To, mCnoIns + mLvl - 1))
                mRge.Insert shift:=Excel.xlShiftToRight
            End If
        End With
    Next
    
    If jj.Crt_Rge_ExtNCol(mRgeBlk, mRgeBlk, 15) Then ss.A 6: GoTo E
    mRgeBlk.Sort Key1:=pRge(1, 1), Order1:=xlAscending, Header:=xlNo _
        , MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
End If

For iRno = pRge.Row To mRnoLas
    mLvl = mWs.Cells(iRno, mCnoOL).Value
    If 1 <= mLvl And mLvl <= 7 Then mWs.Rows(iRno).OutlineLevel = mLvl + 1
    If 8 <= mLvl And mLvl <= 15 Then mWs.Rows(iRno).OutlineLevel = 8
Next

With mWs
    With .OutLine
        .SummaryRow = xlSummaryAbove
        .ShowLevels 2
    End With
    .Activate
    .Application.ActiveWindow.Zoom = 85
End With
pRge(1, pCithOL).EntireColumn.ColumnWidth = 5
Exit Function
R: ss.R
E: Fmt_WsOL_ByCol = True: ss.B cSub, cMod, "pRge,pCithOL,pCithIns", jj.ToStr_Rge(pRge), pCithOL, pCithIns
End Function
Function Fmt_WsOL_ByCol_Tst() As Boolean
Dim mFx$: mFx = "c:\tmp\aa.xls"
Dim mCase As Byte: mCase = 1
Dim mAdr$
Select Case mCase
Case 1
    If jj.Crt_Tbl_ParChd_Tst Then Stop: GoTo E
    If jj.Exp_SetNmtq2Xls("[#]Tmp", mFx, True) Then Stop: GoTo E
    mAdr = "A2"
Case 2
    If False Then
        If jj.Run_Qry("qryFmtWsOL") Then Stop: GoTo E
    End If
    If jj.Exp_SetNmtq2Xls("@[#]MBM3", mFx, True) Then Stop: GoTo E
    mAdr = "A2"
End Select
Dim mWb As Workbook: If jj.Opn_Wb(mWb, mFx, , , True) Then Stop: GoTo E
Dim mWs As Worksheet: Set mWs = mWb.Sheets(1)
If jj.Fmt_WsOL_ByCol(mWs.Range(mAdr), 2, 4) Then Stop: GoTo E
mWb.Application.Visible = True
Stop
GoTo X
E: Fmt_WsOL_ByCol_Tst = True
X: jj.Cls_Wb mWb, False, True
End Function
Function Fmt_WsOL(pWs As Worksheet, pUpToLvl As Byte) As Boolean
'Aim: Use column 1-{pUpToLvl} to set the outline level of {pWs}
Const cSub$ = "Fmt_WsOL"
Dim iRno&, iCno As Byte
For iRno = 1 To 65536
    Dim mIsSet As Boolean:   mIsSet = False ' Is row set the level? If not set, then exit the iRno loop
    For iCno = 1 To pUpToLvl
        If Not IsEmpty(pWs.Cells(iRno, iCno)) Then
            If iCno > 1 Then pWs.Rows(iRno).OutlineLevel = iCno
            mIsSet = True
            Exit For
        End If
    Next
    If Not mIsSet Then Exit For
Next
With pWs
    With .OutLine
        .SummaryRow = xlSummaryAbove
        .ShowLevels 2
    End With
    .Range("$1:$" & pUpToLvl - 1).ColumnWidth = 5
    .Activate
    .Application.ActiveWindow.Zoom = 85
End With
End Function
Function Fmt_yMmmWww(pDte As Date) As String
If pDte < Date Then
    Fmt_yMmmWww = " Past"
    Exit Function
End If
Fmt_yMmmWww = Right(Year(pDte), 1) & "M" & Format(Month(pDte), "00") & "W" & Format(MGIWeekNum(pDte), "00")
End Function
