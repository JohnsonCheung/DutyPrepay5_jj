Attribute VB_Name = "xWsD"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xWsD"
Dim xFbDelta$, xAyWsD() As tWsD, xNWsD As Byte
Type tWsD
    NmWs As String
    AyCno() As Byte
    AmFld() As tMap
    VayvDta As Variant '
End Type
Function WsD_Add_DeltaRec(oDelta&, pNmDelta$, pSess&, pDteCrt As Date, Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: Add a record to pDb![>>Delta]
Const cSub$ = "WsD_Add_DeltaRec"
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, "Select * from [>>Delta]", pDb) Then ss.A 1: GoTo E
With mRs
    .AddNew
    oDelta = !Delta
    !NmDelta = pNmDelta
    !Sess = pSess
    !DteCrt = pDteCrt
    !DteArr = Now
    .Update
End With
GoTo X
R: ss.R
E: WsD_Add_DeltaRec = True: ss.B cSub, cMod, "pSess,pDb", pSess, jj.ToStr_Db(pDb)
X: jj.Cls_Rs mRs
End Function
Function WsD_Add_Delta_FmCsv(pFfnCsvDelta$) As Boolean
'Aim: add {pFfnCsvDeltaDetla} to delta database [>>*].  pFfnCsvDelta fmt: Delta_Tbl_000000_20080519_173705.csv
Const cSub$ = "WsD_Add_Delta_FmCsv"
Dim mA$, mNmDelta$, mSess&, mYYYYMMDD$, mHHMMSS$
Dim mDir$, mFnn$, mExt$: If jj.Brk_Ffn_To3Seg(mDir, mFnn, mExt, pFfnCsvDelta) Then ss.A 1: GoTo E
If mExt <> ".csv" Then ss.A 2, "pFfnCsvDelta must be .csv": GoTo E
If jj.Brk_Str_To5Seg(mA, mNmDelta, mSess, mYYYYMMDD, mHHMMSS, mFnn, "_") Then ss.A 2, "pFfnCsvDelta must be Delta_Tbl_000000_20080519_173705.csv": GoTo E
If mA <> "Delta" Then ss.A 3, "pFfnCsvDelta must start with Delta"
Dim mDteCrt As Date: mDteCrt = CDate(Format(mYYYYMMDD, "0000/00/00") & " " & Format(mHHMMSS, "00:00:00"))
Dim mAcs As Access.Application: If jj.Cv_Acs_FmFb(mAcs, xFbDelta) Then ss.A 3: GoTo E
Dim mDb As DAO.Database: Set mDb = mAcs.CurrentDb
Dim mDelta&: If jj.WsD_Add_DeltaRec(mDelta, mNmDelta, mSess, mDteCrt, mDb) Then ss.A 4: GoTo E
Dim mNmt$, mNmt1$: mNmt = ">#" & mNmDelta: mNmt1 = ">>" & mNmDelta
If jj.Run_Sql("Delete * from [" & mNmt & "]", mAcs) Then ss.A 5: GoTo E
mAcs.DoCmd.TransferText acImportDelim, , mNmt, pFfnCsvDelta, True
Dim mSql$: mSql = jj.Fmt_Str("INSERT INTO [>>{0}]" & _
" SELECT *" & _
" FROM (Select {1} As Delta, x.* from [>#{0}] x)", mNmDelta, mDelta)
If jj.Run_Sql(mSql, mAcs) Then ss.A 5: GoTo E
GoTo X
R: ss.R
E: WsD_Add_Delta_FmCsv = True: ss.B cSub, cMod, "pFfnCsvDelta", pFfnCsvDelta
X: Set mDb = Nothing
   If xFbDelta <> "" Then jj.Cls_CurDb mAcs
End Function
Function WsD_Reset(pWb As Workbook, Optional pFbDelta$ = "") As Boolean
xNWsD = 0
WsD_Reset = WsD_Init_ByWb(pWb, pFbDelta)
End Function
Function WsD_Init_ByWb(pWb As Workbook, Optional pFbDelta$ = "") As Boolean
'Aim: create delta table & xAyWsD for each worksheet (A1=Import:...} of {pWb}
'     delta tables to be create in pFbDelta are
'             each ws:   [>>{NmWs}] & [>#{NmWs}]
'             one delta: [>>delta]
Const cSub$ = "WsD_Init_ByWb"
On Error GoTo R
If xNWsD > 0 Then Exit Function
xFbDelta = pFbDelta
If pFbDelta <> "" Then
    If Not jj.IsFfn(pFbDelta, True) Then If jj.Crt_Fb(pFbDelta) Then ss.A 1: GoTo E
End If
Dim mDb As DAO.Database: If jj.Cv_Db_FmFb(mDb, xFbDelta) Then ss.A 2: GoTo E
If Not jj.IsTbl(">>Delta", mDb) Then
    If jj.Crt_Tbl_FmLoFld(">>Delta", "Delta Long,NmDelta Text 50,Sess Long,DteCrt Date,DteArr Date,DteBegProc Date,DteEndProc Date", 1, , mDb) Then ss.A 1: GoTo E
End If

Dim mXls As Excel.Application: Set mXls = pWb.Application
mXls.ScreenUpdating = False
Dim iWs As Worksheet, mA$
For Each iWs In pWb.Worksheets
    Dim mAyCno() As Byte, mAmFld() As tMap
    Dim mRge As Range: Set mRge = iWs.Range("A" & jj.g.cRnoDta)
    If jj.Crt_Tbl_FmImpWs(mRge, , True, mDb) Then ss.A 3: GoTo E
    If jj.Fnd_AyCnoImpFld(mAyCno, mAmFld, mRge) Then ss.A 4: GoTo E
    ReDim Preserve xAyWsD(xNWsD)
    With xAyWsD(xNWsD)
        .NmWs = iWs.Name
        .AyCno = mAyCno
        .AmFld = mAmFld
    End With
    xNWsD = xNWsD + 1
Next
If mA <> "" Then ss.A 1, "Some ws cannot create Delta Tables", , "The Ws with error", mA: GoTo E
Exit Function
R: ss.R
E: WsD_Init_ByWb = True: ss.B cSub, cMod, "pWb,pFbDelta", jj.ToStr_Wb(pWb), pFbDelta
X: If xFbDelta <> "" Then jj.Cls_Db mDb
   On Error Resume Next
   mXls.ScreenUpdating = True
End Function
#If Tst Then
Function WsD_Init_ByWb_Tst() As Boolean
If True Then
    If jj.Cpy_Fil("P:\AppDef_Meta\MetaDb.xls", "c:\tmp\aa.xls", True) Then Stop: GoTo E
End If
Dim mWb As Workbook: If jj.Opn_Wb(mWb, "c:\tmp\aa.xls", , , True) Then Stop: GoTo E
If WsD_Init_ByWb(mWb, "c:\tmp\aa.mdb") Then Stop: GoTo E
Debug.Print WsD_ToStr_AyWsD
Stop
Dim mAcs As Access.Application: Set mAcs = jj.g.gAcs
If jj.Opn_CurDb(mAcs, "c:\tmp\aa.mdb", , , True) Then Stop: GoTo E
Stop
GoTo X
E: WsD_Init_ByWb_Tst = True
X: jj.Cls_Wb mWb, , True
   jj.Cls_CurDb mAcs
   mAcs.Quit
End Function
#End If
Function WsD_Fnd_iWsD(oWsD%, pNmWs$, Optional pAddIfNotFound As Boolean = False) As Boolean
'Aim: Find oWsD% in xAyWsD, if not found create one
Const cSub$ = "WsD_Fnd_iWsD"
For oWsD = 0 To xNWsD - 1
    If xAyWsD(oWsD).NmWs = pNmWs Then Exit Function
Next
If Not pAddIfNotFound Then ss.A 1, "Given pNmWs is not rec in xAyWsD": GoTo E
ReDim Preserve xAyWsD(xNWsD)
xAyWsD(xNWsD).NmWs = pNmWs
oWsD = xNWsD
xNWsD = xNWsD + 1
Exit Function
E: WsD_Fnd_iWsD = True: oWsD = -1: ss.B cSub, cMod, "pNmWs,pAddIfNotFound", pNmWs, pAddIfNotFound
End Function
Function WsD_Crt_Delta_FmRge(pRge As Range) As Boolean
'Aim: Write the delta change @ {pRge} to a Delta Csv file and import in delta tables.
'     Assume there is already before image @ xAyWsD
Const cSub$ = "WsD_Crt_Delta_FmRge"
Dim mSess$: mSess = Format(Date, "yymmdd") & Format(Now, "hh")
Dim mFfn$: mFfn = "c:\tmp\Delta_Tbl_" & mSess & "_" & Format(Now, "yyyymmdd_hhmmss") & ".csv"
If VBA.Dir(mFfn) <> "" Then mFfn = Fnd_NxtFfn(mFfn)
Dim mCsvStr$
If jj.WsD_Bld_CsvStr_ByRge(mCsvStr, pRge) Then ss.A 1: GoTo E
If mCsvStr = "" Then Exit Function
If jj.Exp_Str_ToFfn(mCsvStr, mFfn) Then ss.A 2: GoTo E
If jj.WsD_Add_Delta_FmCsv(mFfn) Then ss.A 3: GoTo E
Exit Function
E: WsD_Crt_Delta_FmRge = True: ss.B cSub, cMod, "pRge", jj.ToStr_Rge(pRge)
End Function
#If Tst Then
Function WsD_Crt_Delta_FmRge_Tst() As Boolean
If True Then If jj.Cpy_Fil("P:\AppDef_Meta\MetaDb.xls", "c:\tmp\aa.xls", True) Then Stop: GoTo E
Dim mWb As Workbook: If jj.Opn_Wb(mWb, "c:\tmp\aa.xls", , , True) Then Stop: GoTo E
If jj.WsD_Init_ByWb(mWb, "c:\tmp\aa.mdb") Then Stop: GoTo E
Dim mWs As Worksheet: Set mWs = mWb.Sheets("Tbl")
mWs.Activate
Dim mRge As Range: Set mRge = mWs.Range("A" & jj.g.cRnoDta)
If WsD_Set_BefChg(mRge) Then ss.A 5: GoTo E
Stop
If WsD_Crt_Delta_FmRge(mRge) Then Stop: GoTo E
Dim mAcs As Access.Application: Set mAcs = jj.g.gAcs
If jj.Opn_CurDb(mAcs, "c:\tmp\aa.mdb", , , True) Then Stop: GoTo E
Stop
GoTo X
E: WsD_Crt_Delta_FmRge_Tst = True
X: jj.Cls_Wb mWb, , True
   jj.Cls_CurDb mAcs
On Error Resume Next
   mAcs.Quit
End Function
#End If
Function WsD_Set_BefChg(pRge As Range _
    , Optional pCithKey As Byte = 2 _
    ) As Boolean
'Aim: For each cell of the data column @ pRge, put all values to xAyWsD
'     Assume there is already setup in xAyWsD
Const cSub$ = "WsD_Set_BefChg"
'
On Error GoTo R
Dim mWs As Worksheet: Set mWs = pRge.Parent
Dim mWsD%: If WsD_Fnd_iWsD(mWsD, mWs.Name) Then ss.A 1: GoTo E
    
If IsEmpty(pRge(1, 1)) Then
    Dim mAyV(): xAyWsD(mWsD).VayvDta = mAyV
    Exit Function
End If

Dim mAyCno() As Byte: mAyCno = xAyWsD(mWsD).AyCno
Dim mNFld As Byte: mNFld = jj.Siz_Ay(mAyCno): If mNFld = 0 Then ss.A 1, "pAyCno cannot be empty": GoTo E

jj.Shw_AllDta mWs
Dim mRnoBeg&: mRnoBeg = pRge.Row
Dim mRnoEnd&: mRnoEnd = pRge(1, pCithKey).End(xlDown).Row

Dim mNRow%: mNRow = jj.Siz_Ay(xAyWsD(mWsD).VayvDta)
If mNRow = 0 Then
    ReDim xAyWsD(mWsD).VayvDta(mRnoEnd - mRnoBeg, mNFld - 1)
End If

Dim J&, I%, iCno As Byte
With xAyWsD(mWsD)
    For I = 0 To mNFld - 1
        iCno = mAyCno(I) - pRge.Column + 1
        For J = 0 To mRnoEnd - mRnoBeg
            .VayvDta(J, I) = pRge(J + 1, iCno).Value
        Next
    Next
End With
Exit Function
R: ss.R
E: WsD_Set_BefChg = True: ss.B cSub, cMod, "pRge,pCithKey", jj.ToStr_Rge(pRge), pCithKey
End Function
#If Tst Then
Function WsD_Set_BefChg_Tst() As Boolean
Dim mVayvDta()
If False Then If jj.Cpy_Fil("P:\AppDef_Meta\MetaDb.xls", "c:\tmp\aa.xls", True) Then Stop: GoTo E
Dim mWb As Workbook: If jj.Opn_Wb_R(mWb, "c:\tmp\aa.xls", , True) Then Stop: GoTo E
Dim mWs As Worksheet: Set mWs = mWb.Sheets("Tbl")
Dim mRge As Range: Set mRge = mWs.Range("A5")
If jj.WsD_Init_ByWb(mWb, "c:\tmp\aa.mdb") Then Stop: GoTo E
If jj.WsD_Set_BefChg(mRge) Then Stop: GoTo E
Stop
GoTo X
E: WsD_Set_BefChg_Tst = True
X: jj.Cls_Wb mWb, False, True
End Function
#End If
Function WsD_ToStr_AyWsD$()
Dim mS$, J%
For J = 0 To xNWsD - 1
    mS = jj.Add_Str(mS, WsD_ToStr_WsD(xAyWsD(J).NmWs), vbLf)
Next
WsD_ToStr_AyWsD = "xAyWsD(" & xNWsD & ")" & vbLf & mS
End Function
Function WsD_ToStr_WsD$(pNmWs$)
Dim mS$, J%, I%
If WsD_Fnd_iWsD(J, pNmWs) Then GoTo E
With xAyWsD(J)
    For I = 0 To jj.Siz_Am(.AmFld) - 1
        mS = jj.Add_Str(mS, .AmFld(I).F1)
    Next
    WsD_ToStr_WsD = .NmWs & ":" & mS
End With
Exit Function
E: WsD_ToStr_WsD = "Err: WsD_ToStr_WsD(" & pNmWs & ")"
End Function
Function WsD_Bld_CsvStr_ByRge(oCsvStr$, pRgeDta As Range) As Boolean
'Aim: Bld oCsvStr for export if any row has chg: {pRgeDta} does not match {mVayvDta}
'                               any row has add: pRgeDta(1,?) is empty
'                               any row has dlt: exist in {mVayvDta} but not in pRgeDta
'     NmFld row : pRgeDta(0,?)
'     Data Range: pRgeDta(1,1), R2=pRgeDta(2,1).End(xlDown).Row, data col @ mAyCno()
'                 Only data col will be written to oCsvStr
'     mNFld     : # of field (data column)
'     oCsvStr   : Row1: TypDelta, <NmFld1>,..,<mNmFld<mNFld>>
'                 Row2: A,       <NewVal1>,  ..,<NewVal<mNFld>>
'                 Row3: B,       <OldVal1>,  ..,<OldVal<mNFld>>
'                 Row4: D,       <OldVal1>,  ..,<OldVal<mNFld>>
'                 Row5: N,       <NewVal1>,  ..,<NewVal<mNFld>>
'                 ...
'                 Return "" if no chg/add/dlt
Const cSub$ = "WsD_Bld_CsvStr_ByRge"
On Error GoTo R
Dim mWs As Worksheet: Set mWs = pRgeDta.Parent
Dim mAyCno() As Byte, mVayvDta
Dim mWsD%: If jj.WsD_Fnd_iWsD(mWsD, mWs.Name) Then ss.A 1: GoTo E
With xAyWsD(mWsD)
    mAyCno = .AyCno
    mVayvDta = .VayvDta
End With
oCsvStr = ""
jj.Shw_AllDta mWs
Dim mRnoLas&: mRnoLas = pRgeDta(1, 2).End(xlDown).Row

Dim mNFld As Byte: mNFld = jj.Siz_Ay(mAyCno): If mNFld = 0 Then ss.A 1, "mAyCno cannot be zero len": GoTo E
Dim mNROld&: mNROld = jj.Siz_Ay(mVayvDta)

ReDim mAyRHere(mNROld - 1) As Boolean 'Is the Row still here in the ws

Dim iCno As Byte, mV
Dim iRno&, iIdx&, J&, I%, mCno As Byte: mCno = pRgeDta.Column
For J = 0 To mRnoLas - pRgeDta.Row
    iRno = pRgeDta.Row + J
    mV = mWs.Cells(iRno, mCno).Value
    If IsEmpty(mV) Then  ' New row: The Cell(?,1) is empty
        oCsvStr = Add_Str(oCsvStr, "N", vbLf)
        For I = 0 To mNFld - 1
            oCsvStr = oCsvStr & "," & jj.Q_V(mWs.Cells(iRno, mAyCno(I)), True)
        Next
    Else
        'Find iIdx pointing mVayvDta(?,0) ---
        Dim mFound As Boolean: mFound = False
        For iIdx = 0 To mNROld - 1
            If mV = mVayvDta(iIdx, 0) Then mFound = True: Exit For
        Next
        If Not mFound Then ss.A 1, jj.Fmt_Str("mV @ ({0},{1}) not in mVayvDta", iRno, mCno), , "mV", mV: GoTo E
        
        '
        mAyRHere(iIdx) = True ' Mark the mVayvDta(?,0) is here in the Ws.
        
        'The Row iRno match the Data @ iIdx
        Dim mIsChg As Boolean: mIsChg = False

        For I = 0 To mNFld - 1
            If mWs.Cells(iRno, mAyCno(I)).Value <> mVayvDta(iIdx, I) Then mIsChg = True: Exit For
        Next
        If mIsChg Then
            oCsvStr = Add_Str(oCsvStr, "A", vbLf)
            For I = 0 To mNFld - 1
                oCsvStr = oCsvStr & "," & jj.Q_V(mWs.Cells(iRno, mAyCno(I)), True)
            Next
            oCsvStr = oCsvStr & vbCrLf & "B"
            For I = 0 To mNFld - 1
                oCsvStr = oCsvStr & "," & jj.Q_V(mVayvDta(iIdx, I), True)
            Next
        End If
    End If
Next
For J = 0 To mNROld - 1
    If Not mAyRHere(J) Then
        oCsvStr = Add_Str(oCsvStr, "D", vbLf)
        For I = 0 To mNFld - 1
            oCsvStr = oCsvStr & "," & jj.Q_V(mVayvDta(J, I), True)
        Next
    End If
Next
If oCsvStr <> "" Then
    ReDim mAnFld$(mNFld - 1)
    For I = 0 To mNFld - 1
        mAnFld(I) = pRgeDta(0, mAyCno(I))
    Next
    oCsvStr = "TypDelta," & Join(mAnFld, ",") & vbCrLf & oCsvStr
End If
Exit Function
R: ss.R
E: WsD_Bld_CsvStr_ByRge = True: ss.B cSub, cMod, "pRgeDta,mAyCno,mVayvDta", jj.ToStr_Rge(pRgeDta), jj.ToStr_AyByt(mAyCno), "mVayvDta(..)"
End Function
#If Tst Then
Function WsD_Bld_CsvStr_ByRge_Tst() As Boolean
Const cNmWs$ = "Tbl"
If jj.Cpy_Fil("P:\AppDef_Meta\MetaDb.xls", "c:\tmp\aa.xls", True) Then Stop: GoTo E
Dim mWb As Workbook: If jj.Opn_Wb_R(mWb, "c:\tmp\aa.xls", , True) Then Stop: GoTo E
If jj.WsD_Init_ByWb(mWb, "c:\tmp\aa.mdb") Then Stop: GoTo E

Dim mWs As Worksheet: Set mWs = mWb.Sheets(cNmWs)
Dim mRge As Range: Set mRge = mWs.Range("A5")
If jj.WsD_Set_BefChg(mRge) Then Stop: GoTo E
Stop
Dim mCsvStr$: If WsD_Bld_CsvStr_ByRge(mCsvStr, mRge) Then Stop: GoTo E
Debug.Print mCsvStr
Stop
GoTo X
E: WsD_Bld_CsvStr_ByRge_Tst = True
X: jj.Cls_Wb mWb, False, True
End Function
#End If

