Attribute VB_Name = "xRead"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xRead"
Function Read_Str_FmFil(oStr$, pFfn$, Optional pDlt_AftRead As Boolean = False) As Boolean
'Aim: read {oStr} from {pFfn}
Const cSub$ = "Read_Str_FmFil"
On Error GoTo R
Dim mF As Byte: If jj.Opn_Fil_ForInput(mF, pFfn) Then GoTo E
Dim mL$
oStr = ""
While Not EOF(mF)
    Line Input #mF, mL
    oStr = oStr & mL & vbCrLf
Wend
Close #mF
If pDlt_AftRead Then jj.Dlt_Fil pFfn
Exit Function
R: ss.R
E: Read_Str_FmFil = True: ss.B cSub, cMod, "pFfn,pDlt_AftRead", pFfn, pDlt_AftRead
End Function
Function Read_Def_FmtTbl(oDef As tFmtTbl_Def, pQt As QueryTable) As Boolean
Const cSub$ = "Def_FmtTbl"
'Aim: Fill in Following
''    IsSepLin As Boolean '   Default is True: to turn off: After Top Right <Tbl>: NoSepLin
''    Nmtq As String
''    Sq As tSq           ' The square holding the definition
''    RgeDta As Range
''    SqDta As tSq
''    NLvl As Byte
''    Qt As QueryTable
''    RgeCno_Lvl1 As tRgeCno
''    RgeCno_Lvl2 As tRgeCno
''    RgeCno_Lvl3 As tRgeCno
''    RgeCno_Lvl4 As tRgeCno
''    AyCnoSubTot() As Byte
''    AyCnoAvg() As Byte
''    AyCnoFormula() As Byte
''    AyFormula() As String
''    AyCnoVLine_Right() As Byte
''    AyCnoVLine_Left() As Byte
''    AyCnoVLine_RightMedium() As Byte
''    AyCnoVLine_LeftMedium() As Byte
''    AyRgeCno_Merge() As tRgeCno
''    CnoCnt As Byte              ' Column with subtotal of count, Key Word = Cnt
Dim mChk As Boolean
mChk = False

Dim mWs As Worksheet: Set mWs = pQt.Parent
With oDef
    'Find 4 oDef.Sq ' The sq holding the def
    If jj.Fnd_FmtDefSq(.Sq, pQt) Then ss.A 1, "Cannot find Def Sq": GoTo E
    
    '0. Check .Sq.(1,2): 'Nmtq={xxx}
    Dim mA$
    mA$ = mWs.Cells(.Sq.r1, .Sq.c1 + 1).Value
    If Left(mA, 5) <> "Nmtq=" Then ss.A 2, "The cell after first <Tbl> must Nmtq=", , "The value of the cell after <Tbl>,The address Def Sq", mA, jj.ToStr_Sq(.Sq): GoTo E
    .Nmtq = mID(mA, 6)
    If pQt.Name <> .Nmtq Then ss.A 1, "pQt.Name <> the name after Nmtq=<...>", , "Nmtq=<...>", .Nmtq: GoTo E

    '0. Check .Sq.(1,3): 'SepLin={Yes|No}
    mA$ = mWs.Cells(.Sq.r1, .Sq.c1 + 2).Value
    If mA = "SepLin=Yes" Then
        .IsSepLin = True
    ElseIf mA = "SepLin=No" Then
        .IsSepLin = False
    Else
        ss.A 3, "The cell after & after first <Tbl> must SepLin=Yes or SepLin=No", , "The value of the cell after & after first <Tbl>,The address Def Sq", mA, jj.ToStr_Sq(.Sq): GoTo E
    End If
    
    '1. Find .RgeDta, .Qt, .SqDta,
    Set .RgeDta = pQt.ResultRange
    Set .Qt = pQt
    If jj.Cv_Rge2Sq(.SqDta, .RgeDta) Then ss.A 1: GoTo E
    .SqDta.c2 = .Sq.c2 - 1

    '2 Find SubTot
    Dim iRno&: For iRno = .Sq.r1 + 1 To .Sq.r2 - 1
        Dim mV: mV = mWs.Cells(iRno, .Sq.c2).Value
        Select Case mV
        Case "Merge"
                Dim mRgeCno As tRgeCno: If jj.Fnd_RgeCno_InRow(mRgeCno, mWs, iRno, .SqDta.c1, .SqDta.c2 - 1) Then ss.A 6: GoTo E
                ReDim Preserve .AyRgeCno_Merge(0 To jj.Siz_AyRgeCno(.AyRgeCno_Merge)) As tRgeCno
                .AyRgeCno_Merge(UBound(.AyRgeCno_Merge)) = mRgeCno
        Case "VLineLeft":        .AyCnoVLine_Left = jj.Fnd_AyCno_XInRow(mWs, iRno, .SqDta.c1, .SqDta.c2)
        Case "VLineRight":       .AyCnoVLine_Right = jj.Fnd_AyCno_XInRow(mWs, iRno, .SqDta.c1, .SqDta.c2)
        Case "VLineLeftMedium":  .AyCnoVLine_LeftMedium = jj.Fnd_AyCno_XInRow(mWs, iRno, .SqDta.c1, .SqDta.c2)
        Case "VLineRightMedium": .AyCnoVLine_RightMedium = jj.Fnd_AyCno_XInRow(mWs, iRno, .SqDta.c1, .SqDta.c2)
        Case "SubTot":           .AyCnoSubTot = jj.Fnd_AyCno_XInRow(mWs, iRno, 1, .SqDta.c2)
        Case "Avg":              .AyCnoAvg = jj.Fnd_AyCno_XInRow(mWs, iRno, 1, .SqDta.c2)
        Case "Formula":          .AyCnoFormula = jj.Fnd_AyCno_XInRow(mWs, iRno, 1, .SqDta.c2)
                                        Dim N%: N = jj.Siz_Ay(.AyCnoFormula)
                                        If N > 0 Then
                                            ReDim .AyFormula(N - 1)
                                            Dim J%: For J = 0 To N - 1
                                                .AyFormula(J) = mWs.Cells(iRno, .AyCnoFormula(J)).Comment.Text
                                            Next
                                        End If
        Case "Lvl1":             If jj.Fnd_RgeCno_InRow(.RgeCno_Lvl1, mWs, iRno, .SqDta.c1, .SqDta.c2) Then ss.A 6: GoTo E Else If .NLvl < 1 Then .NLvl = 1
        Case "Lvl2":             If jj.Fnd_RgeCno_InRow(.RgeCno_Lvl2, mWs, iRno, .SqDta.c1, .SqDta.c2) Then ss.A 7: GoTo E Else If .NLvl < 2 Then .NLvl = 2
        Case "Lvl3":             If jj.Fnd_RgeCno_InRow(.RgeCno_Lvl3, mWs, iRno, .SqDta.c1, .SqDta.c2) Then ss.A 8: GoTo E Else If .NLvl < 3 Then .NLvl = 3
        Case "Lvl4":             If jj.Fnd_RgeCno_InRow(.RgeCno_Lvl4, mWs, iRno, .SqDta.c1, .SqDta.c2) Then ss.A 9: GoTo E Else If .NLvl < 4 Then .NLvl = 4
        Case "Cnt":              .CnoCnt = jj.Fnd_Cno_XInRow(mWs, iRno, , .SqDta.c1, .SqDta.c2)
        Case Else:               ss.xx 10, cSub, cMod, eWarning, "Invalid Token|Valid Tokens are: VLineLeft,VLineRight,VLineLeftMedium,VLineRightMedium,SubTot,Formula,Lvl1,Lvl2,Lvl3,Lvl4", "Adr,Value", jj.Cv_Cno2Col(.SqDta.c2) & iRno, mV
        End Select
    Next
End With
If mChk Then
    With oDef
        jj.Shw_Dbg cSub, cMod, "Fmt.Tbl Definition: Qt", "Nmtq,IsSepLin", .Nmtq, .IsSepLin
        jj.Shw_Dbg cSub, cMod, "Fmt.Tbl Definition: Sq", "Sq (Definition), SqDta, RgeDta", jj.ToStr_Sq(.Sq), jj.ToStr_Sq(.SqDta), jj.ToStr_Rge(.RgeDta)
        jj.Shw_Dbg cSub, cMod, "Fmt.Tbl Definition: RgeCno", "NLvl,RgeCno_Lvl1,RgeCno_Lvl2,RgeCno_Lvl3,RgeCno_Lvl4", .NLvl, jj.ToStr_RgeCno(.RgeCno_Lvl1), jj.ToStr_RgeCno(.RgeCno_Lvl2), jj.ToStr_RgeCno(.RgeCno_Lvl3), jj.ToStr_RgeCno(.RgeCno_Lvl4)
        jj.Shw_Dbg cSub, cMod, "Fmt.Tbl Definition: Ay", "CnoCnt,AyCnoFormula,AyCnoFormula,AyCnoSubTot,AyCnoVLineLeft,AyCnoVLineLeftMedium,AyCnoVLineRight,AyCnoVLineRightMedium", .CnoCnt, jj.Cv_AyByt2LoCol(.AyCnoFormula), jj.ToStr_Ays(.AyFormula, , vbLf), jj.Cv_AyByt2LoCol(.AyCnoSubTot), jj.Cv_AyByt2LoCol(.AyCnoVLine_Left), jj.Cv_AyByt2LoCol(.AyCnoVLine_LeftMedium), jj.Cv_AyByt2LoCol(.AyCnoVLine_Right), jj.Cv_AyByt2LoCol(.AyCnoVLine_RightMedium)
    End With
    Stop
End If
Exit Function
R: ss.R
E: Read_Def_FmtTbl = True: ss.B cSub, cMod, "pQt", jj.ToStr_Qt(pQt)
End Function
Function Read_MacroFil(oAm() As tMap, pFfnMacro$) As Boolean
Const cSub$ = "MacroFil"
Dim mFno As Byte: If jj.Opn_Fil_ForInput(mFno, pFfnMacro) Then ss.A 1: GoTo E
Dim N%
While Not EOF(mFno)
    Dim mLine$: Line Input #mFno, mLine
    If Left(mLine, 1) <> "#" Then
        ReDim Preserve oAm(N)
        If jj.Brk_Str2Map(oAm(N), mLine) Then ss.A 2: GoTo E
    End If
Wend
Close #mFno
Exit Function
R: ss.R
E: Read_MacroFil = True: ss.B cSub, cMod, "pFfnMarcro"
X: Close #mFno
End Function


