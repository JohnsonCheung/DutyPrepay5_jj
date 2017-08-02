Attribute VB_Name = "xExp"
Option Compare Text
#Const Tst = True
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xExp"
Function Exp_TpSrc() As Boolean
Const cSub$ = "Exp_TpSrc"
Dim mF As Byte, mOFil$
mOFil = jj.Sdir_Doc & "Tp_Doc.csv"
If jj.Opn_Fil_ForOutput(mF, mOFil, True) Then ss.A 1: GoTo E
 
Dim mDirTp$: mDirTp = jj.Sdir_Tp
If jj.Exp_TpSrc_InDir(mDirTp, mF) Then ss.A 2: GoTo E

Dim iSubFolder As Folder
For Each iSubFolder In g.gFso.GetFolder(mDirTp).SubFolders
    Dim mDir As String
    mDir = iSubFolder.Name
    If mDir <> "." And mDir <> ".." Then If jj.Exp_TpSrc_InDir(mDirTp & mDir, mF) Then ss.A 3: GoTo E
Next
Close #mF
'Format the csv to xls
Dim mWb As Workbook, mWs As Worksheet
If jj.Opn_Wb_RW(mWb, mOFil) Then ss.A 4: GoTo E
Set mWs = mWb.Worksheets(1)
If jj.Fmt_WsOL(mWs, 3) Then ss.A 5: GoTo E
mWs.Columns(3).ColumnWidth = 40
mWs.Columns(4).ColumnWidth = 15
If jj.Dlt_Fil(Left(mOFil, Len(mOFil) - 4) & ".xls") Then ss.A 5: GoTo E
mWb.SaveAs Left(mOFil, Len(mOFil) - 4) & ".xls", Excel.XlFileFormat.xlWorkbookNormal
mWb.Application.Visible = True
Exit Function
R: ss.R
E: Exp_TpSrc = True: ss.B cSub, cMod
End Function
Function Exp_TpSrc_InDir(pDirTp$, pF As Byte) As Boolean
'Aim: Exp all the datasource of all xls files in {pDir} to <pF>
Const cSub$ = "Exp_TpSrc_InDir"
'==Start==
Dim mAyFn$(): If jj.Fnd_AyFn(mAyFn, pDirTp) Then ss.A 1: GoTo E
Dim J%
For J = 0 To jj.Siz_Ay(mAyFn) - 1
    Dim mWb As Workbook: If jj.Opn_Wb_R(mWb, pDirTp & mAyFn(J)) Then ss.A 1: GoTo E
    With mWb
        Write #pF, mWb.Name, , , , mWb.FullName
        If mWb.PivotCaches.Count > 0 Then
            Write #pF, , "PivotCaches.Count(" & mWb.PivotCaches.Count & ")"
            Dim iPc As Excel.PivotCache
            For Each iPc In .PivotCaches
                Write #pF, , , iPc.CommandText, , iPc.Connection
            Next
        End If
        Dim iWs As Worksheet
        For Each iWs In mWb.Worksheets
            If iWs.PivotTables.Count > 0 Then
                Write #pF, , "PivotTables.Count(" & iWs.PivotTables.Count & ") Ws(" & iWs.Name & ")"
                Dim iPt As PivotTable
                For Each iPt In iWs.PivotTables
                    Write #pF, , , iPt.PivotCache.CommandText, iPt.Name, iPt.PivotCache.Connection
                Next
            End If
        Next
        For Each iWs In mWb.Worksheets
            If iWs.QueryTables.Count > 0 Then
                Write #pF, , "QueryTables.Count(" & iWs.QueryTables.Count & ") Ws(" & iWs.Name & ")"
                Dim iQt As Excel.QueryTable
                For Each iQt In iWs.QueryTables
                    Write #pF, , , iQt.CommandText, iQt.Name, iQt.Connection
                Next
            End If
        Next
        .Close False
    End With
Next
Exit Function
R: ss.R
E: Exp_TpSrc_InDir = True: ss.B cSub, cMod, "pDirTp", pDirTp
End Function
Function Exp_Str_ToFfn(pStr$, pFfn$, Optional pOvrWrt As Boolean = False) As Boolean
'Aim: export {pStr} to {pFfn} with {pOvrWrt} optional
Const cSub$ = "Exp_Str_ToFfn"
Dim mF As Byte: If jj.Opn_Fil_ForOutput(mF, pFfn, pOvrWrt) Then GoTo E
Print #mF, pStr
Close #mF
Exit Function
R: ss.R
E: Exp_Str_ToFfn = True: ss.B cSub, cMod, "pStr,pFfn,pOvrWrt", pStr, pFfn, pOvrWrt
End Function
Function Exp_Tbl2FfnXml(pNmt$, pFfnXml$) As Boolean
Const cSub$ = "Exp_Tbl2FfnXml"
Dim mNmt0$: mNmt0 = jj.Rmv_SqBkt(pNmt)
Dim mFlag As Access.AcExportXMLOtherFlags
mFlag = acEmbedSchema
Application.ExportXML acExportTable, mNmt0, pFfnXml, , , , , mFlag
Exit Function
R: ss.R
E: Exp_Tbl2FfnXml = True: ss.B cSub, cMod, "pNmt,pFfnXml", pNmt, pFfnXml
End Function
#If Tst Then
Function Exp_Tbl2FfnXml_Tst() As Boolean
If jj.Exp_Tbl2FfnXml("mstBrand", "c:\tmp\mstBrand.xml") Then Stop
End Function
#End If
Function Exp_Rf(pFfn$, Optional pNmPrj$ = "", Optional pAcs As Access.Application = Nothing) As Boolean
Const cSub$ = "Exp_Rf"
On Error GoTo R
Dim mPrj As VBProject: If jj.Cv_Prj(mPrj, pNmPrj, pAcs) Then ss.A 1: GoTo E
Dim iRf As VBIDE.Reference
Dim mFno As Byte: If jj.Opn_Fil_ForOutput(mFno, pFfn, True) Then ss.A 2: GoTo E
For Each iRf In mPrj.References
    With iRf
        Write #mFno, .Name, .FullPath, .BuiltIn, .Type
    End With
Next
GoTo X
Exit Function
R: ss.R
E: Exp_Rf = True: ss.B cSub, cMod, "pNmPrj", pNmPrj
X:
    Close #mFno
End Function
#If Tst Then
Function Exp_Rf_Tst() As Boolean
If jj.Exp_Rf("c:\tmp\aa.txt") Then Stop
Shell "notepad c:\tmp\aa.txt", vbMaximizedFocus
End Function
#End If
Function Exp_Pgm(pFxTar$ _
    , Optional pFbSrc$ = "" _
    , Optional pFxTp$ = "" _
    ) As Boolean
'Aim: Export all pgm in {pFbSrc} to {pFxTar}!@OldPgm,@OldArg.  If {pFxTar} not exist, copy from {pFxTp}.
'     If no ws OldQry in pFxTar, create a OldQry at End, otherwise, the OldQry is expect to have an Export Format.
Const cSub$ = "Exp_Pgm"
Const cLnt$ = "@OldPgm,@OldArg"
If jj.Exp_Pgm_ToTbl(cLnt, pFbSrc) Then ss.A 1: GoTo E
If jj.Exp_SetNmtq2Xls_wFmt(cLnt, pFxTar, pFxTp) Then ss.A 2: GoTo E
GoTo X
E: Exp_Pgm = True: ss.B cSub, cMod, "pFxTar,pFbSrc,pFxTp", pFxTar, pFbSrc, pFxTp
X:
End Function
Function Exp_Pgm_ToTbl( _
      Optional pLnt$ = "#OldPgm:#OldArg" _
    , Optional pFbSrc$ = "" _
    ) As Boolean
'Aim: Export all pgm in {pFbSrc} to #OldPgm,#OldArg (defined in {pLnt}).
Const cSub$ = "Exp_Pgm_ToTbl"
On Error GoTo R
Dim mAcs As Access.Application: If jj.Cv_Acs_FmFb(mAcs, pFbSrc) Then ss.A 1: GoTo E
Dim mAnPrj$(), mNPrj%
Do
    If jj.Fnd_AnPrj(mAnPrj, , , mAcs) Then ss.A 2: GoTo E
    mNPrj = jj.Siz_Ay(mAnPrj)
    If mNPrj = 0 Then ss.A 3, "No Prj is found in pFbSrc": GoTo E
Loop Until True

Dim mNmtOldPgm$, mNmtOldArg$
Do
    Dim mA$(): mA = Split(Replace(pLnt, ":", cComma), cComma)
    If jj.Siz_Ay(mA) <> 2 Then ss.A 4, "pLnt must have 2 elements": GoTo E
    mNmtOldPgm = Trim(mA(0))
    mNmtOldArg = Trim(mA(1))
    Dim mDPgm As New d_Pgm: If mDPgm.CrtTbl(mNmtOldPgm) Then ss.A 5: GoTo E
    Dim mDArg As New d_Arg: If mDArg.CrtTbl(mNmtOldArg) Then ss.A 6: GoTo E
    Dim mRsPgm As DAO.Recordset: Set mRsPgm = CurrentDb.TableDefs(mNmtOldPgm).OpenRecordset
    Dim mRsArg As DAO.Recordset: Set mRsArg = CurrentDb.TableDefs(mNmtOldArg).OpenRecordset
Loop Until True

Dim iPrj%
For iPrj = 0 To mNPrj - 1
    Dim mNmPrj$: mNmPrj = mAnPrj(iPrj)
    Dim mPrj As VBProject: If jj.Fnd_Prj(mPrj, mNmPrj, mAcs) Then ss.A 8: GoTo E
    Dim mAnm$(): If jj.Fnd_Anm_ByPrj(mAnm, mPrj) Then ss.A 9: GoTo E
    
    Dim iMd%
    For iMd = 0 To jj.Siz_Ay(mAnm) - 1
        Dim mNmm$: mNmm = mAnm(iMd)
        jj.Shw_Sts "Export pgm in module [" & mNmPrj & "." & mNmm & "]...."
        Dim mMd As CodeModule: If jj.Fnd_Md(mMd, mPrj, mNmm) Then ss.A 10: GoTo E
        Dim mAnPrc$(): If jj.Fnd_AnPrc_ByMd(mAnPrc, mMd, , True) Then ss.A 11: GoTo E
        Dim iPrc%
        For iPrc = 0 To jj.Siz_Ay(mAnPrc) - 1
            Dim mNmPrc$: mNmPrc = mAnPrc(iPrc):
            If mNmPrc = "qBrkRec" And mNmm = "Brk" Then Stop
            'Debug.Print "Prj(" & iPrj & ":" & mNmPrj & ") Md(" & iMd & ":" & mNmm & ") Prc(" & iPrc & ":" & mNmPrc & ")"
            Dim mPrcBody$: If jj.Fnd_PrcBody_ByMd(mPrcBody, mMd, mNmPrc, True) Then ss.A 12: GoTo E
            Dim mAyDArg() As d_Arg
            If jj.Brk_PrcBody(mDPgm, mAyDArg, mPrcBody) Then ss.A 13: GoTo E
            mDPgm.x_NmPrj = mNmPrj
            mDPgm.x_Nmm = mNmm
            If mDPgm.Ins(mRsPgm) Then ss.A 14: GoTo E
            If mDArg.InsAy(mRsArg, mNmPrj, mNmm, mNmPrc, mAyDArg) Then ss.A 16: GoTo E
        Next
    Next
Next
GoTo X
R: ss.R
E: Exp_Pgm_ToTbl = True: ss.B cSub, cMod, "pLnt,pFbSrc", pLnt, pFbSrc
X:
    If pFbSrc <> "" Then jj.Cls_CurDb mAcs
    jj.Cls_Rs mRsPgm
    jj.Cls_Rs mRsArg
    jj.Clr_Sts
End Function
#If Tst Then
Function Exp_Pgm_ToTbl_Tst() As Boolean
Const cSub$ = "Exp_Pgm_ToTbl_Tst"
Dim mLnt$, mFbSrc$
Dim mCase As Byte: mCase = 1
Select Case mCase
Case 1
    mLnt$ = "#OldPgm,#OldArg"
    mFbSrc = "p:\workingdir\pgmobj\JMtcLgc.Mdb"
Case 2
    mLnt$ = "#OldPgm,#OldArg"
    mFbSrc = "p:\workingdir\pgmobj\JMtcDb.Mdb"
End Select
If jj.Exp_Pgm_ToTbl(mLnt, mFbSrc) Then Stop: GoTo E
DoCmd.OpenTable "#OldPgm"
DoCmd.OpenTable "#OldArg"
GoTo X
E: Exp_Pgm_ToTbl_Tst = True
X:
End Function
#End If
Function Exp_Qry(pFxTar$ _
    , Optional pExclQDpd As Boolean = False _
    , Optional pFbSrc$ = "" _
    , Optional pFxTp$ = "" _
    ) As Boolean
'Aim: Export all query in {pFbSrc} to {pFxTar}!OldQry.  If {pFxTar} not exist, copy from {pFxTp}.
'     If no ws OldQry in pFxTar, create a OldQry at End, otherwise, the OldQry is expect to have an Export Format.
Const cSub$ = "Exp_Qry"
Const cLnt$ = "@OldQry,@OldQsT"
If jj.Exp_Qry_ToTbl(cLnt, pExclQDpd, pFbSrc) Then ss.A 1: GoTo E
If jj.Crt_Qry("qryTmp_Oup_OldQry", "Select * from [@OldQry] order by Len(Sql) Desc") Then ss.A 2: GoTo E
If jj.Crt_Qry("qryTmp_Oup_OldQsT", "Select * from [@OldQsT] order by Len(LnFld) Desc") Then ss.A 3: GoTo E
If jj.Exp_SetNmtq2Xls_wFmt("qryTmp_Oup_OldQry,qryTmp_Oup_OldQsT", pFxTar, pFxTp) Then ss.A 4: GoTo E
If jj.Dlt_Qry_ByPfx("qryTmp_Oup_OldQ") Then ss.A 5: GoTo E
GoTo X
E: Exp_Qry = True: ss.B cSub, cMod, "pFxTar,pExclQDpd,pFbSrc,pFxTp", pFxTar, pExclQDpd, pFbSrc, pFxTp
X: jj.Dlt_Qry_ByPfx "qryTmp_Oup_Old"
End Function
#If Tst Then
Function Exp_Qry_Tst() As Boolean
Const cSub$ = "Exp_Qry_Tst"
Dim mFbSrc$: mFbSrc = "P:\WorkingDir\PgmObj\JMtcDb.mdb"
Dim mFx$:
Dim mInclQDpd As Boolean
Dim mFxTp$
Dim mWb As Workbook
Dim mCase As Byte: mCase = 2
Select Case mCase
Case 1
    mFx = "C:\tmp\bb.csv"
    If jj.Exp_Qry_ToCsv(mFx, mInclQDpd, mFbSrc) Then ss.A 1: GoTo E
Case 2
    mFx = "P:\AppDef_Meta\MetaLgc.xls"
    mFxTp = ""
    mInclQDpd = True
    If jj.Exp_Qry(mFx, mInclQDpd, mFbSrc, mFxTp) Then ss.A 2: GoTo E
End Select
If jj.Opn_Wb_RW(mWb, mFx) Then ss.A 3: GoTo E
mWb.Application.Visible = True
Exit Function
R: ss.R
E: Exp_Qry_Tst = True: ss.B cSub, cMod
X: jj.Cls_Wb mWb, True, True
End Function
#End If
Function Exp_Qry_ToTbl( _
      Optional pLnt$ = "#OldQry:#OldQsT" _
    , Optional pInclQDpd As Boolean = False _
    , Optional pFbQry$ = "" _
    ) As Boolean
'Aim: Export all query in {pFbQry} to #OldQry (defined in {pLnt}).
'     If {pInclQDpd}, #OldQsT will also be included.  To get #OldQsT, it will jj.Run_Qry to each [mAnQs] in {pFbQry}
Const cSub$ = "Exp_Qry_ToTbl"

Dim mAcs As Access.Application: If jj.Cv_Acs_FmFb(mAcs, pFbQry) Then ss.A 1: GoTo E

Dim mAnQs$(), mNQs%
Do
    Dim mDbQry As DAO.Database: Set mDbQry = mAcs.CurrentDb
    If jj.Fnd_AnQs(mAnQs, , mDbQry) Then ss.A 2: GoTo E
    mNQs = jj.Siz_Ay(mAnQs) - 1
    If mNQs = 0 Then ss.A 3, "No QrySet is found": GoTo E
Loop Until True

Dim mNmtOldQry$, mNmtOldQsT$
Do
    Dim mA$(): mA = Split(Replace(pLnt, ":", cComma), cComma)
    Select Case jj.Siz_Ay(mA)
    Case 1
        mNmtOldQry = Trim(mA(0))
        If pInclQDpd Then ss.A 4, "pInclQDpd is True and pLnt should have 2 elements": GoTo E
    Case 2
        mNmtOldQry = Trim(mA(0))
        mNmtOldQsT = Trim(mA(1))
    Case Else
        ss.A 5, "pLnt must have 1 or 2 element": GoTo E
    End Select
    Dim mDQry As New d_Qry: If mDQry.CrtTbl(mNmtOldQry) Then ss.A 6: GoTo E
    Dim mLgcT As New d_QsT: If mLgcT.CrtTbl(mNmtOldQsT) Then ss.A 7: GoTo E
Loop Until True

Dim iQs%
For iQs = 0 To mNQs - 1
    If jj.Exp_Qs_ToTbl(mAnQs(iQs), pLnt, pInclQDpd, mAcs) Then ss.A 8: GoTo E
Next
GoTo X
R: ss.R
E: Exp_Qry_ToTbl = True: ss.B cSub, cMod, "pLnt,pInclQDpd,pFbQry", pLnt, pInclQDpd, pFbQry
X: If pFbQry <> "" Then jj.Cls_CurDb mAcs
   Set mDbQry = Nothing
End Function
#If Tst Then
Function Exp_Qry_ToTbl_Tst() As Boolean
Const cFb$ = "P:\WorkingDir\PgmObj\JMtcDb.mdb"
If jj.Exp_Qry_ToTbl(, True, cFb) Then Stop: GoTo E
DoCmd.OpenTable "#OldQry"
DoCmd.OpenTable "#OldQsT"
Exit Function
E:
End Function
#End If
Function Exp_Qs_ToTbl(pNmQs$ _
    , Optional pLnt$ = "#OldQry:#OldQsT" _
    , Optional pInclQDpd As Boolean = False _
    , Optional pAcs As Access.Application = Nothing _
    ) As Boolean
'Aim: Export all queries of {pFbSrc}!{pNmQs} to {pNmt} in currentdb with option of {pInclQDpd}
Const cSub$ = "Exp_Qs_ToTbl"
jj.Shw_Sts "Exporting DQry of Qs[" & pNmQs & "] with pInclQDpd[" & pInclQDpd & "]...."
Dim mNmtOldQry$, mNmtOldQsT$
Do
    Dim mA$(): mA = Split(Replace(pLnt, ":", cComma), cComma)
    Select Case jj.Siz_Ay(mA)
    Case 1
        mNmtOldQry = Trim(mA(0))
        If pInclQDpd Then ss.A 1, "pInclQDpd is True and pLnt should have 2 elements": GoTo E
    Case 2
        mNmtOldQry = Trim(mA(0))
        mNmtOldQsT = Trim(mA(1))
    Case Else
        ss.A 2, "pLnt must have 1 or 2 element": GoTo E
    End Select
Loop Until True

Dim mAcs As Access.Application: Set mAcs = jj.Cv_Acs(pAcs)
Dim mDbQry As DAO.Database: Set mDbQry = mAcs.CurrentDb
Dim mFb$: mFb = mDbQry.Name

If pInclQDpd Then If jj.Run_Qry(pNmQs, , , True, , , , mAcs) Then ss.A 3: GoTo E
Dim mAyDQry() As jj.d_Qry, NDQry%
If jj.Fnd_AyDQry(mAyDQry, pNmQs & "_*", pInclQDpd, mAcs) Then ss.A 4: GoTo E

Dim mNDQry%: mNDQry = jj.Siz_AyDQry(mAyDQry)
If mNDQry > 0 Then If mAyDQry(0).InsAy(mAyDQry, mFb, mNmtOldQry) Then ss.A 5: GoTo E

If Not pInclQDpd Then GoTo X
Dim mAyOldQsT() As d_QsT, mNOldQsT%, iDQry%
mNOldQsT = -1
For iDQry = 0 To mNDQry - 1
    Dim mAnt$(): mAnt = Split(mAyDQry(iDQry).LnTbl, cComma)
    Dim mNTbl%: mNTbl = jj.Siz_Ay(mAnt)
    Dim iTbl%
    For iTbl = 0 To mNTbl - 1
        Dim iOldQsT%, mNmt$
        mNmt = Trim(mAnt(iTbl))
        For iOldQsT = 0 To mNOldQsT
            If mAyOldQsT(iOldQsT).x_NmTbl = mNmt Then GoTo Nxt_Tbl
        Next
        mNOldQsT = mNOldQsT + 1
        ReDim Preserve mAyOldQsT(mNOldQsT)
        Set mAyOldQsT(mNOldQsT) = New d_QsT
        With mAyOldQsT(mNOldQsT)
            .x_Fb = mFb
            .x_NmQs = pNmQs
            .x_NmTbl = mNmt
            .x_LnFld = jj.ToStr_Nmt(mNmt, True, , , , mDbQry)
        End With
Nxt_Tbl:
    Next
Next

'Write to #OldQsT
If mNOldQsT > 0 Then If mAyOldQsT(0).InsAy(mAyOldQsT, mNmtOldQsT) Then ss.A 8: GoTo E
GoTo X
R: ss.R
E: Exp_Qs_ToTbl = True: ss.B cSub, cMod, "pNmQs,pLnt,pInclQDpd,pAcs", pNmQs, pLnt, jj.ToStr_Acs(pAcs), pInclQDpd
X: jj.Clr_Sts
   Set mDbQry = Nothing
End Function
#If Tst Then
Function Exp_Qs_ToTbl_Tst() As Boolean
Dim mFbSrc$
Dim mNmQs$
Dim mInclQDpd As Boolean
Dim mCase As Byte: mCase = 2
Select Case mCase
Case 1
    mFbSrc = "p:\workingdir\pgmobj\JMtcLgc.Mdb"
    mNmQs = "qryExpLgc"
    mInclQDpd = True
Case 2
    mFbSrc = "p:\workingdir\pgmobj\JMtcDb.Mdb"
    mNmQs = "qryAddTblR"
    mInclQDpd = True
End Select
Dim mAcs As Access.Application: Set mAcs = jj.g.gAcs: If jj.Opn_CurDb(mAcs, mFbSrc) Then Stop: GoTo E
Dim mDQry As New d_Qry: If mDQry.CrtTbl Then Stop: GoTo E
Dim mOldQsT As New d_QsT: If mOldQsT.CrtTbl Then Stop: GoTo E
If jj.Exp_Qs_ToTbl(mNmQs, , mInclQDpd, mAcs) Then Stop: GoTo E
DoCmd.OpenTable "#OldQsT"
DoCmd.OpenTable "#OldQry"
GoTo X
E: Exp_Qs_ToTbl_Tst = True
X: If mFbSrc <> "" Then jj.Cls_CurDb mAcs
End Function
#End If
Function Exp_Qry_ToCsv(pFc$, Optional pWrtHdr As Boolean = False, Optional pFbSrc$ = "") As Boolean
'Aim: Export all query in {pFbSrc} to {pFc} start at {pRnoBeg}.  If {pRnoBeg}=1, put the header
Const cSub$ = "Exp_Qry_ToCsv"
Dim mAyDQry() As jj.d_Qry
Dim mDb As DAO.Database: If jj.Cv_Db_FmFb(mDb, pFbSrc) Then ss.A 1: GoTo E
If jj.Fnd_AyDQry(mAyDQry, "qry*", True, mDb) Then ss.A 3: GoTo E
Dim N%: N = jj.Siz_AyDQry(mAyDQry): If N = 0 Then ss.A 3: GoTo E

Dim mF As Byte: If jj.Opn_Fil_ForOutput(mF, pFc, True) Then ss.A 1: GoTo E
If pWrtHdr Then
    Write #mF, "NmMdb";
    If mAyDQry(0).WrtHdr(mF) Then ss.A 4: GoTo E
End If

Dim mFb$: mFb = mDb.Name
Dim J%
For J = 0 To N - 1
    If mAyDQry(J).Wrt(mF, mFb) Then ss.A 5: GoTo E
Next
GoTo X
E: Exp_Qry_ToCsv = True: ss.B cSub, cMod, "pFc,pWrtHdr,pFbSrc", pFc, pWrtHdr, pFbSrc
X: If pFbSrc <> "" Then jj.Cls_Db mDb
   Close mF
End Function
Function Exp_Qry_ToWs(pWs As Worksheet, Optional oRnoBeg& = 1, Optional pFbSrc$ = "", Optional pInclQDpd As Boolean = False) As Boolean
'Aim: Export all query in {pFbSrc} to {pWs} start at {pRnoBeg}.  If {pRnoBeg}=1, put the header
Const cSub$ = "Exp_Qry_ToWs"
Dim mAyDQry() As jj.d_Qry
Dim mDb As DAO.Database: If jj.Cv_Db_FmFb(mDb, pFbSrc) Then ss.A 1: GoTo E
If jj.Fnd_AyDQry(mAyDQry, "qry*", pInclQDpd, mDb) Then ss.A 3: GoTo E
Dim N%: N = jj.Siz_AyDQry(mAyDQry): If N = 0 Then ss.A 3: GoTo E

If oRnoBeg = 1 Then If jj.Set_Ws_ByLv(pWs, 1, 1, False, "NmMdb,NmQs,Maj,Min,Rest,Typ,Sql,LnQDpd") Then ss.xx 3, cSub, cMod: GoTo E
Dim mNmMdb$: mNmMdb = mDb.Name
Dim J%
For J = 0 To N - 1
    oRnoBeg = oRnoBeg + 1
    With pWs
        .Cells(oRnoBeg, 1).Value = mNmMdb
        .Cells(oRnoBeg, 2).Value = mAyDQry(J).NmQs
        .Cells(oRnoBeg, 3).Value = mAyDQry(J).Maj
        .Cells(oRnoBeg, 4).Value = mAyDQry(J).Min
        .Cells(oRnoBeg, 5).Value = mAyDQry(J).Rest
        .Cells(oRnoBeg, 6).Value = jj.ToStr_TypQry(mAyDQry(J).Typ)
        .Cells(oRnoBeg, 7).Value = mAyDQry(J).Sql
        .Cells(oRnoBeg, 8).Value = mAyDQry(J).LnTbl
    End With
Next
GoTo X
E: Exp_Qry_ToWs = True: ss.B cSub, cMod, "pWs,oRnoBeg,pFbSrc", jj.ToStr_Ws(pWs), oRnoBeg, pFbSrc
X: If pFbSrc <> "" Then jj.Cls_Db mDb
End Function
Function Exp_SetNmtq2Mdb_ByTblOup(pFbTar$) As Boolean
'Aim: In CurrentDb, export tables as defined in [tblOup] to {pFbTar} (Create if not exist).  If no [tblOup], no export.
Const cSub$ = "Exp_SetNmtq2Mdb_ByTblOup"
If pFbTar = "" Then ss.A 1: GoTo E
If Not jj.IsTbl("tblOup") Then Exit Function
If VBA.Dir(pFbTar) = "" Then If jj.Crt_Fb(pFbTar) Then ss.A 1: GoTo E
With CurrentDb.TableDefs("tblOup").OpenRecordset
    While Not .EOF
        If InStr(!LikNmt, "*") > 0 Then
            If jj.Snd_Tbl_ToMdb(!LikNmt, pFbTar$) Then ss.A 2: GoTo E
        Else
            Dim mSql$: mSql = jj.Fmt_Str("Select * into {0} in '{1}' from {0}", !LikNmt, pFbTar)
            If jj.Run_Sql(mSql) Then ss.A 5: GoTo E
        End If
        .MoveNext
    Wend
    .Close
End With
Exit Function
R: ss.R
E: Exp_SetNmtq2Mdb_ByTblOup = True: ss.B cSub, cMod, "pFbTar", pFbTar
End Function
#If Tst Then
Function Exp_SetNmtq2Mdb_ByTblOup_Tst() As Boolean
If jj.Exp_SetNmtq2Mdb_ByTblOup("C:\aa.mdb") Then Stop
End Function
#End If
Function Exp_Nmtq2Mdb(pNmtq$, pFbTar$, Optional pNmtTar$ = "", Optional pFbSrc$ = "", Optional pOvrWrt As Boolean = False) As Boolean
'Aim: Export {pNmtq} in {p.FbSrc} to table {p.NmtTar} in {p.FbTar}.  {Nmt2Mdb} will be created if not exist
Const cSub$ = "Exp_Nmq2Mdb"
On Error GoTo R
If VBA.Dir(pFbTar) = "" Then If jj.Crt_Fb(pFbTar) Then ss.A 1: GoTo E
Dim mNmtTar$: mNmtTar = IIf(pNmtTar = "", pNmtq, pNmtTar)
On Error GoTo R
Dim mIn_FbSrc$: If pFbSrc <> "" Then mIn_FbSrc = " in '" & pFbSrc & cQSng
Dim mSql$: mSql = jj.Fmt_Str("select * into {0} in '{1}' from {2}{3}", mNmtTar, pFbTar, pNmtq, mIn_FbSrc)
If jj.Run_Sql(mSql) Then ss.A 2: GoTo E
Exit Function
R: ss.R
E: Exp_Nmtq2Mdb = True: ss.B cSub, cMod, "pNmq,pFbTar,pNmtTar", pNmtq, pFbTar, pNmtTar
End Function
#If Tst Then
Function Exp_Nmtq2Mdb_Tst() As Boolean
Dim mCase As Byte, mNmtq$, mFbTar1$, mFbTar2$, mNmtTar1$, mNmtTar2$
mFbTar1 = "c:\aa.mdb"
mFbTar2 = "c:\bb.mdb"
If jj.Dlt_Fil(mFbTar1) Then Stop
If jj.Dlt_Fil(mFbTar2) Then Stop
For mCase = 1 To 2
    Select Case mCase
    Case 1: mNmtq = "qryAllBrand"
    Case 2: mNmtq = "query1"
    End Select
    If jj.Exp_Nmtq2Mdb(mNmtq, mFbTar1) Then Stop
Next
If jj.Exp_Nmtq2Mdb("qryAllBrand", mFbTar2, mNmtTar1) Then Stop
If jj.Exp_Nmtq2Mdb("query1", mFbTar2, mNmtTar2) Then Stop

g.gAcs.OpenCurrentDatabase mFbTar1
g.gAcs.Visible = True
Dim mAcs As New Access.Application
mAcs.OpenCurrentDatabase mFbTar2
mAcs.Visible = True
Stop
End Function
#End If
Function Exp_SetNmtq2Xls_wFmt(pSetNmtq$, pFxTar$ _
    , Optional pFxTp$ = "" _
    , Optional pNmWsPfx$ = "", Optional pNmWsSfx$ = "" _
    , Optional pFbSrc$ = "" _
    , Optional pNoExpTim As Boolean = False _
    ) As Boolean
'Aim: Export all tables/queries in {pSetNmtq} to {pFxTar} with {pNmWsPfx/pNmWsSfx} added to each Ws (ie ws name will be pPfx + Nmtq + pSfx}.
'     "Note to Nmtq": if Nmtq is in format of xxx_Oup_yyy or #@yyy, yyy will be use
Const cSub$ = "Exp_SetNmtq2Xls_wFmt"
If VBA.Dir(pFxTar$) = "" Then
    If VBA.Dir(pFxTp$) <> "" Then If jj.Cpy_Fil(pFxTp, pFxTar) Then ss.A 1: GoTo E
End If
On Error GoTo R
Dim mAntq$(): If jj.Fnd_Antq_BySetNmtq(mAntq, pSetNmtq) Then ss.A 2: GoTo E

Dim mWb As Workbook, mToBeDelete$: mToBeDelete = ""
If VBA.Dir(pFxTar) = "" Then
    mToBeDelete = "ToBeDelete": If jj.Crt_Wb(mWb, pFxTar, , mToBeDelete) Then ss.A 1: GoTo E
Else
    If jj.Opn_Wb_RW(mWb, pFxTar) Then ss.A 1: GoTo E
End If

Dim mAntqErr$()
Dim I%: For I = 0 To jj.Siz_Ay(mAntq) - 1
    Dim mWs As Worksheet
    Dim mNmWsTar$: mNmWsTar = pNmWsPfx & jj.Cut_Aft(jj.Cut_Aft(jj.Cut_Aft(mAntq(I), "_Oup_"), "#@"), "@") & pNmWsSfx
    If jj.Fnd_Ws(mWs, mWb, mNmWsTar, True) Then
        If jj.Add_Ws(mWs, mWb, mNmWsTar) Then jj.Add_AyEle mAntqErr, mAntq(I): GoTo Nxt
        If jj.Exp_Nmtq2Ws(mAntq(I), mWs, pFbSrc) Then jj.Add_AyEle mAntqErr, mAntq(I): GoTo Nxt
    Else
        If jj.Exp_Nmtq2Ws_wFmt_ByCpyRs(mAntq(I), mWs.Range("A5"), pFbSrc, pNoExpTim) Then jj.Add_AyEle mAntqErr, mAntq(I): GoTo Nxt
    End If
Nxt:
Next
If mToBeDelete$ <> "" Then jj.Dlt_Ws_InWb mWb, mToBeDelete
If jj.Siz_Ay(mAntqErr) > 0 Then ss.A 3, "These tables cannot be exported: " & Join(mAntqErr, ","): GoTo E
If jj.Cls_Wb(mWb, True) Then ss.A 4: GoTo E
Exit Function
R: ss.R
E: Exp_SetNmtq2Xls_wFmt = True: ss.B cSub, cMod, "pSetNmtq, pFxTar, pNmWsPfx, pNmWsSfx", pSetNmtq, pFxTar, pNmWsPfx, pNmWsSfx
End Function
#If Tst Then
Function Exp_SetNmtq2Xls_wFmt_Tst() As Boolean
Const cSub$ = "Exp_SetNmtq2Xls_wFmt_Tst"
Dim mSetNmtq$, mFxTar$, mNmWsPfx$, mNmWsSfx$, mFbSrc$, mFxTp$, mNoExpTim As Boolean
Dim mCase As Byte: mCase = 2
Select Case mCase
Case 1
    mSetNmtq = "[#]@Ele"
    mFxTar = "c:\tmp\ExpMetaDb_All.xls"
    mFxTp = "P:\AppDef_Meta\MetaDb.xls"
Case 2
    'If jj.Crt_Qry("qryTmp_Oup_OldQry", "Select * from [@OldQry] order by Len(Sql) Desc") Then ss.a 2: GoTo E
    mFbSrc = "p:\workingdir\pgmobj\JMtcLgc.Mdb"
    Dim mDb As DAO.Database: If jj.Opn_Db_RW(mDb, mFbSrc) Then ss.A 4: GoTo E
    If jj.Crt_Qry("qryTmp_Oup_OldQsT", "Select * from [@OldQsT] order by Len(LnFld) Desc", mDb) Then ss.A 3: GoTo E
    mSetNmtq = "qryTmp_Oup_OldQsT" 'qryTmp_Oup_OldQry,qryTmp_Oup_OldQsT
    mFxTar = "P:\AppDef_Meta\MetaLgc.xls"
    mFxTp = ""
End Select
If jj.Exp_SetNmtq2Xls_wFmt(mSetNmtq, mFxTar, mFxTp, mNmWsPfx, mNmWsSfx, mFbSrc, mNoExpTim) Then Stop
Dim mWb As Workbook: If jj.Opn_Wb_RW(mWb, mFxTar, True, True) Then Stop
Stop
GoTo X
E: Exp_SetNmtq2Xls_wFmt_Tst = True: ss.B cSub, cMod
X:
    jj.Dlt_Qry_ByPfx "qryTmp_Oup_OldQ", mDb
    jj.Cls_Db mDb
End Function
#End If
Function Exp_SetNmtq2Xls(pSetNmtq$, pFxTar$, Optional pOvrWrt As Boolean = False, Optional pNmWsPfx$ = "", Optional pNmWsSfx$ = "", Optional pFbSrc$ = "") As Boolean
'Aim: Export all tables/queries in {pSetNmtq} to {pFxTar} with {pNmWsPfx/pNmWsSfx} added to each Ws (ie ws name will be pPfx + Nmtq + pSfx}.
'     "Note to Nmtq": if Nmtq is in format of xxx_Oup_yyy, yyy will be use
Const cSub$ = "Exp_SetNmtq2Xls"
If jj.Ovr_Wrt(pFxTar, pOvrWrt) Then ss.A 1: GoTo E
On Error GoTo R
Dim mA$
Dim mAntq$(): If jj.Fnd_Antq_ByLik(mAntq, pSetNmtq) Then ss.A 2: GoTo E
Dim I%: For I = 0 To jj.Siz_Ay(mAntq) - 1
    Dim mNmtq$, p%: p = InStr(mAntq(I), "_Oup_")
    If p > 0 Then
        mNmtq = mID(mAntq(I), p + 5)
    Else
        mNmtq = mAntq(I)
    End If
    
    If jj.Exp_Nmtq2Xls(mAntq(I), pFxTar, pNmWsPfx & mNmtq & pNmWsSfx, pFbSrc) Then mA = jj.Add_Str(mA, mAntq(I))
Next
If mA <> "" Then ss.A 2, "These tables cannot be exported: " & mA: GoTo E
Exit Function
R: ss.R
E: Exp_SetNmtq2Xls = True: ss.B cSub, cMod, "pSetNmtq, pFxTar, pNmWsPfx, pNmWsSfx", pSetNmtq, pFxTar, pNmWsPfx, pNmWsSfx
End Function
#If Tst Then
Function Exp_SetNmtq2Xls_Tst() As Boolean
Const cSub$ = "Exp_SetNmtq2Xls_Tst"
Dim mSetNmtq$
Dim mFx$
Dim mOvrWrt As Boolean
Dim mNmWsPfx$
Dim mNmWsSfx$
Dim mCase As Byte, mRslt As Boolean
mCase = 1
Select Case mCase
Case 1
    mSetNmtq = "mst*,tbl*,tmp*"
    mFx = "C:\Tmp\export.xls"
    mOvrWrt = True
    mNmWsPfx = "sss_"
    mNmWsSfx = "_xxx"
Case 2
    mSetNmtq = "[#]*"
    mFx = "P:\WorkingDir\Export\ExpDir.xls"
    mOvrWrt = True
    mNmWsPfx = ""
    mNmWsSfx = ""
End Select
mRslt = jj.Exp_SetNmtq2Xls(mSetNmtq, mFx, mOvrWrt, mNmWsPfx, mNmWsSfx)
jj.Shw_Dbg cSub, cMod, , "mRslt,mSetNmtq,mFx,mNmWsSfx", mRslt, mSetNmtq, mFx, mNmWsSfx
Dim mWb As Workbook: If jj.Opn_Wb(mWb, mFx, , , True) Then Stop: GoTo E
Exit Function
E: Exp_SetNmtq2Xls_Tst = True
End Function
#End If
Function Exp_Nmtq2Ws_wFmt_ByQt(pNmtq$, pRge As Range _
    , Optional pFbSrc$ = "" _
    , Optional pNoExpTim As Boolean = False _
    ) As Boolean
'Aim: Read data from table {pFbSrc}!{pNmQt} to pNmQt.Destination
Const cSub$ = "Exp_Nmtq2Ws_wFmt_ByQt"
On Error GoTo R
Dim mWs As Worksheet: Set mWs = pRge.Parent

'Build Qt
Clr_Qt mWs
Dim mFbSrc$
If pFbSrc = "" Then
    mFbSrc = CurrentDb.Name
Else
    mFbSrc = pFbSrc
End If
Dim mQt As QueryTable: Set mQt = mWs.QueryTables.Add(jj.CnnStr_Mdb(mFbSrc), pRge)
With mQt
    Dim mSql$: If jj.BldSql_Qt(mSql, pRge, pNmtq, pFbSrc) Then ss.A 2: GoTo E
    .CommandType = xlCmdSql
    .CommandText = mSql
    .BackgroundQuery = False
    .AdjustColumnWidth = False
    .FillAdjacentFormulas = False
    .MaintainConnection = False
    .PreserveColumnInfo = True
    .PreserveFormatting = True
    .FieldNames = False
End With

'Fill Data
Shw_AllDta mWs
Dim mAdr$: mAdr = pRge.Address
mWs.Range(mWs.Cells(pRge.Row, 1), mWs.Cells(65536, 1)).EntireRow.ClearFormats
mWs.Range(mWs.Cells(pRge.Row, 1), mWs.Cells(65536, 1)).EntireRow.Delete
Set pRge = mWs.Range(mAdr)

If jj.Rfh_Qt(mQt) Then ss.A 3: GoTo E
Dim mNRec&: mNRec = mQt.ResultRange.Rows.Count

'Fmt Qt
If jj.Fmt_Ws(pRge, mNRec, 3) Then ss.A 4: GoTo E

Clr_Qt mWs
Exit Function
R: ss.R
E: Exp_Nmtq2Ws_wFmt_ByQt = True: ss.B cSub, cMod, "pNmtq,pWs,pFbSrc,pNoExpTim", pNmtq, jj.ToStr_Rge(pRge), pFbSrc, pNoExpTim
End Function
Function Exp_Nmtq2Ws_wFmt_ByCpyRs(pNmtq$, pRge As Range _
    , Optional pFbSrc$ = "" _
    , Optional pNoExpTim As Boolean = False _
    ) As Boolean
'Aim: Export {pNmtq} in {pFbSrc} to {pWs}
'     If {pNpWsTar} exist,
'       it must be in following format
'       - #1: A1      : A1 must be in format of import:<mNpWsTar>
'       - #2: Rno4    : It the rno of field name.  All vbstring fields will be used as field name.
'                     - The cmt of the field name is used as Rno5's formula and will copy to all records
'       - #3: Rno5    : Is the data rno
'       - #4: SubTot  : Col B is be subtotal to count the # of records
'       otherwise, error
Const cSub$ = "Exp_Nmtq2Ws_wFmt_ByCpyRs"
On Error GoTo R
Dim mStp!
jj.Shw_Sts "Formatting Ws[" & pRge.Parent.Name & "] ..."
mStp = 1 'Check A1

'Build mSql
Dim mSql$: If jj.BldSql_Qt(mSql, pRge, pNmtq, pFbSrc) Then ss.A 1: GoTo E

Dim mIsMem255 As Boolean
If jj.IfMem255(mIsMem255, pNmtq, pFbSrc) Then ss.A 2: GoTo E
'Fill Data
mStp = 3
Dim mWs As Worksheet: Set mWs = pRge.Parent
Shw_AllDta mWs
Dim mAdr$: mAdr = pRge.Address
mWs.Range(mWs.Cells(pRge.Row, 1), mWs.Cells(65536, 1)).EntireRow.ClearFormats
mWs.Range(mWs.Cells(pRge.Row, 1), mWs.Cells(65536, 1)).EntireRow.Delete
Set pRge = mWs.Range(mAdr)

Dim mNRec&
If mIsMem255 Then
    If jj.Cpy_FmRs(mNRec, pRge, mSql) Then ss.A 13: GoTo E
Else
    Dim mRs As DAO.Recordset
    If jj.Opn_Rs(mRs, mSql) Then ss.A 12: GoTo E
    mNRec = mWs.Range("A5").CopyFromRecordset(mRs)
    mRs.Close
End If

'Fmt Data
If jj.Fmt_Ws(pRge, mNRec) Then ss.A 13: GoTo E
GoTo X
R: ss.R
E: Exp_Nmtq2Ws_wFmt_ByCpyRs = True: ss.B cSub, cMod, "pNmtq,pWs,pFbSrc,mStp", pNmtq, jj.ToStr_Rge(pRge), pFbSrc, mStp
X: jj.Cls_Rs mRs
   jj.Clr_Sts
End Function
#If Tst Then
Function Exp_Nmtq2Ws_wFmt_ByQt_Tst() As Boolean
Exp_Nmtq2Ws_wFmt_ByQt_Tst = Exp_Nmtq2Ws_wFmt_Tst(False)
End Function
#End If
#If Tst Then
Function Exp_Nmtq2Ws_wFmt_ByCpyRs_Tst() As Boolean
Exp_Nmtq2Ws_wFmt_ByCpyRs_Tst = Exp_Nmtq2Ws_wFmt_Tst(True)
End Function
#End If
#If Tst Then
Function Exp_Nmtq2Ws_wFmt_Tst(pByCpyRs As Boolean) As Boolean
Const cSub$ = "Exp_Nmtq2Ws_wFmt_Tst"
Dim mNmtq$, mFxTar$, mFbSrc$, mNmWs$
Dim mWb As Workbook, mWs As Worksheet
Dim mCase As Byte: mCase = 5
Select Case mCase
Case 1
    mNmtq = "$Stp"
    mFxTar = "c:\book1.xls"
    mNmWs = "Stp"
    mFbSrc = ""
    Const cLv$ = "Stp,NmStp,NmLgc,StpNo,NmtTar,OldQsTTar,IsOup,LnFld,NQry,TotQry,LnStpNoChd,Des,IsOpt"
    Const cSqlIns$ = "Insert into Stp (" & cLv & ") values"
    If True Then
        If jj.Crt_Tbl_FmLoFld(mNmtq, "Stp Int, NmStp Text 10, NmLgc Text 10, StpNo Byte, NmtTar Text 10, OldQsTTar Text 10, IsOup Text 1, LnFld Text 255, NQry Int, TotQry Int, LnStpNoChd Text 255, Des Memo, IsOpt YesNo", 1) Then Stop
        '                         Stp, NmStp   , NmLgc, StpNo, NmtTar, OldQsTTar, IsOup, LnFld     , NQry, TotQry, LnStpNoChd, Des
        If jj.Run_Sql(cSqlIns & "(1  , 'bb_01' , 'bb' , 1    , 'xxx' , 'xxx'  , 'Y'  , 'xx,xx,bb', 4   , 6     , '12,33'   , 'xxxxx', True)") Then Stop: GoTo E
        If jj.Run_Sql(cSqlIns & "(2  , 'bb_02' , 'bb' , 2    , 'xxx' , 'xxx'  , Null , 'xx,xx,bb', 4   , 6     , '12,33'   , 'xxxxx', False)") Then Stop: GoTo E
        If jj.Run_Sql(cSqlIns & "(3  , 'bb_03' , 'bb' , 3    , 'xxx' , 'xxx'  , 'Y'  , 'xx,xx,bb', 4   , 6     , '12,33'   , 'xxxxx', True)") Then Stop: GoTo E
        If jj.Run_Sql(cSqlIns & "(4  , 'bb_04' , 'bb' , 4    , 'xxx' , 'xxx'  , Null , 'xx,xx,bb', 4   , 6     , '12,33'   , 'xxxxx', True)") Then Stop: GoTo E
        
        If jj.Crt_Wb(mWb, mFxTar, True, mNmWs) Then Stop: GoTo E
        Set mWs = mWb.Sheets(1)
        If jj.Set_Ws_ByLv(mWs, 4, 1, False, cLv) Then Stop: GoTo E
        mWs.Range("A1").Value = "Import:" & mNmWs
        jj.Cls_Wb mWb, True
        mFbSrc = "p:\WorkingDir\MetaAll.Mdb"
    End If
Case 2
    mNmtq = "@Schm"
    mFxTar = "P:\AppDef_Meta\MetaDb.xls"
    mNmWs = "Schm"
    mFbSrc = "P:\WorkingDir\PgmObj\JMtcDb.mdb"
Case 3
    mNmtq = "@TblUF"
    mFxTar = "P:\AppDef_Meta\MetaDb.xls"
    mNmWs = "TblUF"
    mFbSrc = "P:\WorkingDir\PgmObj\JMtcDb.mdb"
Case 4
    mNmWs = "Stp    "
    mNmtq = "@Stp"
    mFxTar = "c:\tmp\aa.xls"
    mNmWs = "QryT"
    mFbSrc = "p:\workingdir\pgmobj\JMtcLgc.mdb"
    If jj.Cpy_Fil("P:\AppDef_Meta\MetaLgc.xls", mFxTar, True) Then Stop: GoTo E
Case 5
    mNmtq = "@Dir"
    mFxTar = "P:\AppDef_Meta\MetaDb.xls"
    mNmWs = "Dir"
    mFbSrc = "P:\WorkingDir\PgmObj\JMtcDb.mdb"
End Select
If jj.Opn_Wb_RW(mWb, mFxTar, , True) Then Stop: GoTo E
Set mWs = mWb.Sheets(mNmWs)
Dim mRge As Range: Set mRge = mWs.Range("A5")
If pByCpyRs Then
    If jj.Exp_Nmtq2Ws_wFmt_ByCpyRs(mNmtq, mRge, mFbSrc) Then Stop: GoTo E
Else
    If jj.Exp_Nmtq2Ws_wFmt_ByQt(mNmtq, mRge, mFbSrc) Then Stop: GoTo E
End If
If jj.Shw_Dbg(cSub, cMod, "mNmtq,mFxTar,mNmWs,mFbSrc", mNmtq, mFxTar, mNmWs, mFbSrc) Then Stop: GoTo E
Stop
GoTo X
Exit Function
E: Exp_Nmtq2Ws_wFmt_Tst = True
X: jj.Cls_Wb mWb, True, True
End Function
#End If
Function Exp_Nmtq2Xls_wFmt(pNmtq$, pFxTar$ _
    , Optional pNmWsTar$ = "" _
    , Optional pFbSrc$ = "" _
    , Optional pNoExpTim As Boolean = False _
    ) As Boolean
'Aim: Export {pNmtq} in {pFbSrc} to {pFxTar}!{pNmWsTar}.
'     If {pNmWsTar} exist,
'       it must be in following format
'       - #1: A1      : A1 must be in format of import:<mNmWsTar>
'       - #2: Rno4    : It the rno of field name.  All vbstring fields will be used as field name.
'                     - The cmt of the field name is used as Rno5's formula and will copy to all records
'       - #3: Rno5    : Is the data rno
'       - #4: SubTot  : Col B is be subtotal to count the # of records
'       otherwise, error
'     Else
'       Export without format
'     endif
'
'Logic:
'   Opn [mWs]
'   Chk A1
'   Build mRs
'   CopyRecordSet
'   CopyFormula
'   SubTot
Const cSub$ = "Exp_Nmtq2Xls_wFmt"
'On Error Goto R
'
'Dim mNmtq$: mNmtq = jj.Rmv_SqBkt(pNmtq)
''Open mWb & mWs
'Dim mWb As Workbook, mWs As Worksheet
'Do
'    'FxTar no exist, create one & Export without fmt
'    Dim mNmWsTar$: mNmWsTar = jj.NonBlank(pNmWsTar, mNmtq)
'    If VBA.Dir(pFxTar) = "" Then
'        If jj.Crt_Fx(pFxTar, , mNmWsTar) Then ss.a 1: GoTo E
'        If jj.Exp_Nmtq2Xls(mNmtq, pFxTar, mNmWsTar, pFbSrc) Then ss.a 3: GoTo E
'        Exit Function
'    End If
'
'    'FxTar exist, but ws not: create one & Export with fmt
'    If jj.Opn_Wb_RW(mWb, pFxTar) Then ss.a 1: GoTo E
'    If Not jj.IsWs(mWb, mNmWsTar) Then
'        If jj.Cls_Wb(mWb, False) Then ss.a 2: GoTo E
'        If jj.Exp_Nmtq2Xls(mNmtq, pFxTar, mNmWsTar, pFbSrc) Then ss.a 3: GoTo E
'        Exit Function
'    End If
'
'    'FxTar & ws exist, set the mWs
'    Set mWs = mWb.Sheets(mNmWsTar)
'Loop Until True
'
'If Nmtq2Ws_wFmt(pNmtq, mWs, pFbSrc, pNoExpTim) Then ss.a 4: GoTo E
'If jj.Cls_Wb(mWb, True) Then ss.a 5: GoTo E
'Goto X
'R: ss.R
'E: Exp_Nmtq2Xls_wFmt = True: ss.B cSub, cMod, "pNmtq,pFxTar,pNmWsTar,pFbSrc", pNmtq, pFxTar, pNmWsTar, pFbSrc
'X:
'    jj.Cls_Wb mWb, False, True
End Function
#If Tst Then
Function Exp_Nmtq2Xls_wFmt_Tst() As Boolean
Const cSub$ = "Exp_Nmtq2Xls_wFmt_Tst"
Dim mNmtq$, mFxTar$, mNmWsTar$, mFbSrc$
Dim mCase As Byte: mCase = 4
Select Case mCase
Case 1
    mNmtq = "Stp"
    mFxTar = "c:\book1.xls"
    mNmWsTar = ""
    mFbSrc = ""
    Const cSqlIns$ = "Insert into Stp (Stp, NmStp, NmLgc, StpNo, NmtTar, OldQsTTar, IsOup, LnFld, NQry, TotQry, LnStpNoChd, Des, IsOpt) values"
    If True Then
        If jj.Crt_Tbl_FmLoFld(mNmtq, "Stp Int, NmStp Text 10, NmLgc Text 10, StpNo Byte, NmtTar Text 10, OldQsTTar Text 10, IsOup Text 1, LnFld Text 255, NQry Int, TotQry Int, LnStpNoChd Text 255, Des Memo, IsOpt YesNo", 1) Then Stop
        '                         Stp, NmStp   , NmLgc, StpNo, NmtTar, OldQsTTar, IsOup, LnFld     , NQry, TotQry, LnStpNoChd, Des
        If jj.Run_Sql(cSqlIns & "(1  , 'bb_01' , 'bb' , 1    , 'xxx' , 'xxx'  , 'Y'  , 'xx,xx,bb', 4   , 6     , '12,33'   , 'xxxxx', True)") Then Stop: GoTo E
        If jj.Run_Sql(cSqlIns & "(2  , 'bb_02' , 'bb' , 2    , 'xxx' , 'xxx'  , Null , 'xx,xx,bb', 4   , 6     , '12,33'   , 'xxxxx', False)") Then Stop: GoTo E
        If jj.Run_Sql(cSqlIns & "(3  , 'bb_03' , 'bb' , 3    , 'xxx' , 'xxx'  , 'Y'  , 'xx,xx,bb', 4   , 6     , '12,33'   , 'xxxxx', True)") Then Stop: GoTo E
        If jj.Run_Sql(cSqlIns & "(4  , 'bb_04' , 'bb' , 4    , 'xxx' , 'xxx'  , Null , 'xx,xx,bb', 4   , 6     , '12,33'   , 'xxxxx', True)") Then Stop: GoTo E
    End If
Case 2
    mNmtq = "tmpExpTbl_Oup_Tbl"
    mFxTar = "P:\WorkingDir\Export\MetaDb.xls"
    mNmWsTar = "Tbl"
    mFbSrc = ""
Case 3
    mNmtq = "[#@Ele]"
    mFxTar = "c:\tmp\ExpMetaDb_All.xls"
    mNmWsTar = "Ele"
    mFbSrc = ""
Case 4
    mNmtq = "@Stp"
    mFxTar = "c:\tmp\aa.xls"
    mNmWsTar = "QryT"
    mFbSrc = "p:\workingdir\pgmobj\JMtcLgc.mdb"
    If jj.Cpy_Fil("P:\AppDef_Meta\MetaLgc.xls", mFxTar, True) Then Stop: GoTo E
End Select
If jj.Exp_Nmtq2Xls_wFmt(mNmtq, mFxTar, mNmWsTar, mFbSrc) Then Stop: GoTo E
If jj.Shw_Dbg(cSub, cMod, "mNmtq, mFxTar, mNmWsTar, mFbSrc", mNmtq, mFxTar, mNmWsTar, mFbSrc) Then Stop: GoTo E
Dim mWb As Workbook: If jj.Opn_Wb_RW(mWb, mFxTar, True, True) Then Stop: GoTo E
Exit Function
E: Exp_Nmtq2Xls_wFmt_Tst = True
End Function
#End If
Function Exp_Sql2Xls(pSql$, pFxTar$, pNmWsTar$) As Boolean
'Aim: Export the result of the {pSql} to {pFxTar}!{pNmWsTar}
'     Note: if {pNmWsTar} exist in {pFxTar}, {pNmWsTar} will be replace and position of the ws in pFxTar will be retended.
'     Note: if {pNmWsTar} not exist in {pFxTar}, {pNmWsTar} will be added at end.
Const cSub$ = "Exp_Sql2Xls"
Const cSql$ = "SELECT * INTO [Excel 8.0;HDR=YES;IMEX=2;DATABASE={0}].[{1}] FROM ({2})"

Dim mIsWs_InFx As Boolean
If VBA.Dir(pFxTar) = "" Then
    mIsWs_InFx = False
Else
    If jj.IfWs_InFx(mIsWs_InFx, pFxTar, pNmWsTar) Then ss.A 1: GoTo E
End If

Dim mSql$
If mIsWs_InFx Then
    Dim mNmWsTar_New$: mNmWsTar_New = Left(pNmWsTar, 31 - 4) & "_New"
    mSql = jj.Fmt_Str(cSql, pFxTar, mNmWsTar_New, pSql)
    If jj.Run_Sql(mSql) Then ss.A 2: GoTo E
    If jj.Repl_Ws_InFx(pFxTar, pNmWsTar, mNmWsTar_New) Then ss.A 3: GoTo E
Else
    mSql = jj.Fmt_Str(cSql, pFxTar, pNmWsTar, pSql)
    If jj.Run_Sql(mSql) Then ss.A 4: GoTo E
End If
Exit Function
R: ss.R
E: Exp_Sql2Xls = True: ss.B cSub, cMod, "pSql,pFxTar,pNmWsTar", pSql, pFxTar, pNmWsTar
End Function
Function Exp_Nmtq2Ws(pNmtq$, pWs As Worksheet _
    , Optional pFbSrc$ = "" _
    ) As Boolean
Stop
End Function
Function Exp_Nmtq2Xls(pNmtq$, pFxTar$ _
    , Optional pNmWsTar$ = "" _
    , Optional pFbSrc$ = "" _
    ) As Boolean
'Aim: Export {pNmtq} in {pFbSrc} to {pFxTar}!{pNmWsTar} by "select * into {pFxTar}"
'     Note: if {pNmWsTar} exist in {pFxTar}, {pNmWsTar} will be replace and position of the ws in pFxTar will be retended.
'     Note: if {pNmWsTar} not exist in {pFxTar}, {pNmWsTar} will be added at end.
Const cSub$ = "Exp_Nmtq2Xls"
Const cSql$ = "SELECT * INTO [Excel 8.0;HDR=YES;IMEX=2;DATABASE={0}].[{1}] FROM [{2}]{3}"
Dim mNmWsTar$: mNmWsTar = Left(Fct.NonBlank(pNmWsTar, pNmtq), 31)
Dim mInFbSrc$: mInFbSrc = jj.Cv_Fb2InFb(pFbSrc)
Dim mNmtq0$: mNmtq0 = jj.Rmv_SqBkt(pNmtq)
Dim mSql$:   mSql = jj.Fmt_Str(cSql, pFxTar, mNmWsTar, mNmtq0, mInFbSrc)
ss.xMonMsg = 3010
If jj.Run_Sql(mSql) Then
    If Not ss.xMonMsgMatched Then ss.A 1: GoTo E
    
    Dim mNmWsTar_New$: mNmWsTar_New = Left(mNmWsTar, 31 - 4) & "_New"
    mSql = jj.Fmt_Str(cSql, pFxTar, mNmWsTar_New, mNmtq0, mInFbSrc)
    If jj.Run_Sql(mSql) Then ss.A 2: GoTo E
    If jj.Repl_Ws_InFx(pFxTar, mNmWsTar, mNmWsTar_New) Then ss.A 3: GoTo E
End If
Exit Function
R: ss.R
E: Exp_Nmtq2Xls = True: ss.B cSub, cMod, "pNmtq,pFxTar,pNmWsTar,pFbSrc", pNmtq, pFxTar, pNmWsTar, pFbSrc
End Function
Function Exp_Nmtq2Xls_Tst() As Boolean
Const cSub$ = "Exp_Nmtq2Xls_Tst"
Dim mNmtq$:     mNmtq = "mstBrand"
Dim mFx$:       mFx = "c:\tmp\aa.xls"
Dim mNmWs$:     mNmWs = "xxx"
Dim mFbSrc$:    mFbSrc = ""
Dim mRslt As Boolean
If jj.Exp_Nmtq2Xls(mNmtq, mFx, "xxx") Then Stop
Stop
Dim mWb As Workbook, mWs As Worksheet
If jj.Crt_Wb(mWb, mFx, True, mNmWs) Then Stop
Set mWs = mWb.Sheets(mNmWs)
If jj.Set_Ws_ByLpAp(mWs, 1, "xx,bb,dd", 1, 4, 3) Then Stop
If jj.Cls_Wb(mWb, True) Then Stop
If jj.Exp_Nmtq2Xls(mNmtq, mFx, mNmWs, mFbSrc) Then Stop
If jj.Exp_Nmtq2Xls(mNmtq, mFx, mNmWs & "1", mFbSrc) Then Stop
jj.Shw_Dbg cSub, cMod, , "mRslt, mNmtq, mFx, mNmWs, mFbSrc", mRslt, mNmtq, mFx, mNmWs, mFbSrc
If jj.Opn_Wb_R(mWb, mFx, , True) Then Stop
End Function
Function Exp_SetNmtq2Mdb(pSetNmtq$, pFbTar$, Optional pPfx$ = "", Optional pSfx$ = "", Optional pFbSrc$ = "") As Boolean
'Aim: Export all tables of {pSetNmtq$} in {pFbSrc} to {pFbTar} with table name having pPfx/pSfx added
Const cSub$ = "Exp_SetNmtq2Mdb"
Dim mDbSrc As DAO.Database: If jj.Cv_Db_FmFb(mDbSrc, pFbSrc) Then ss.A 1: GoTo E
On Error GoTo R
Dim mAyLikNmtq$(): mAyLikNmtq = Split(pSetNmtq$, cComma)
Dim mA$, J%, I%
For J = 0 To jj.Siz_Ay(mAyLikNmtq) - 1
    Dim mAnt$(): If jj.Fnd_Ant_ByLik(mAnt, mAyLikNmtq(J), mDbSrc) Then ss.A 2: GoTo E
    For I = 0 To jj.Siz_Ay(mAnt) - 1
        If jj.Exp_Nmtq2Mdb(mAnt(I), pFbTar, pPfx & mAnt(I) & pSfx, pFbSrc$) Then mA = jj.Add_Str(mA, mAnt(J))
    Next
    '
    Dim mAnq$(): If jj.Fnd_Anq_ByLik(mAnq, mAyLikNmtq(J), mDbSrc) Then ss.A 3: GoTo E
    For I = 0 To jj.Siz_Ay(mAnq) - 1
        If jj.Exp_Nmtq2Mdb(mAnq(I), pFbTar, pPfx & mAnq(I) & pSfx, pFbSrc$) Then mA = jj.Add_Str(mA, mAnq(J))
    Next
Next
If mA <> "" Then ss.xx 5, cSub, cMod, eRunTimErr, "Some table cannot be exported", "List of Tables cannot exported,pSetNmtq,pFbTar", mA, pSetNmtq, pFbTar: GoTo E
GoTo X
R: ss.R
E: Exp_SetNmtq2Mdb = True: ss.B cSub, cMod, "J,mAntq(J),pSetNmtq,pPfx,pSfx", J, mAyLikNmtq(J), pSetNmtq, pPfx, pSfx
X: jj.Cls_Db mDbSrc
End Function
#If Tst Then
Function Exp_SetNmtq2Mdb_Tst() As Boolean
Dim mFbTar1$: mFbTar1 = "c:\aa.mdb"
Dim mFbTar2$: mFbTar2 = "c:\bb.mdb"
If jj.Dlt_Fil(mFbTar1) Then Stop
If jj.Exp_SetNmtq2Mdb("qry*,tbl*,mst*", mFbTar1, "Wrk_") Then Stop
If jj.Exp_SetNmtq2Mdb("Wrk*", mFbTar2, , , mFbTar1) Then Stop
g.gAcs.OpenCurrentDatabase mFbTar1
g.gAcs.Visible = True
Dim mAcs As New Access.Application
mAcs.OpenCurrentDatabase mFbTar2
mAcs.Visible = True
Stop
End Function
#End If
Function Exp_SetNmtq2Dir(pSetNmtq$, pDir$, Optional pPfx$ = "", Optional pSfx$ = "") As Boolean
'Aim: Export all tables in {pSetNmtq$} to Excel in {pDir} with {pPfx/pSfx} added.  Ws Nam will use the table name
Const cSub$ = "Exp_SetNmtq2Dir"
On Error GoTo R
Dim mAyLikNmtq$(): mAyLikNmtq = Split(pSetNmtq$, cComma)
Dim mA$
Dim J%: For J = 0 To jj.Siz_Ay(mAyLikNmtq) - 1
    Dim mAntq$(): If jj.Fnd_Antq_ByLik(mAntq, mAyLikNmtq(J)) Then ss.A 1: GoTo E
    Dim I%: For I = 0 To jj.Siz_Ay(mAntq) - 1
        If jj.Exp_Nmtq2Xls(mAntq(I), pDir & pPfx & mAntq(I) & pSfx & ".xls") Then mA = jj.Add_Str(mA, mAntq(J))
    Next
Next
If mA <> "" Then ss.A 1, "Some table cannot be exported", , "List of Tables cannot exported,pSetNmtq,pDir", mA, pSetNmtq, pDir: GoTo E
Exit Function
R: ss.R
E: Exp_SetNmtq2Dir = True: ss.B cSub, cMod, "J,mAntq(J),pSetNmtq,pDir,pSfx", J, mAntq(J), pSetNmtq, pDir, pSfx
End Function
Function Exp_SetNmtq2Dir_Tst() As Boolean
Const cSub$ = "Exp_SetNmtq2Dir_Tst"
Dim mDir$: mDir = "C:\Tmp\"
Dim mPfx$: mPfx = "sss_"
Dim mSfx$: mSfx = "_xxx"
Dim mSetNmtq$: mSetNmtq = "mst*,tbl*,tmp*"
Dim mCase As Byte, mRslt As Boolean
mCase = 1
Select Case mCase
Case 1
    mRslt = jj.Exp_SetNmtq2Dir(mSetNmtq, mDir, mPfx, mSfx)
Case 2
    mRslt = jj.Exp_SetNmtq2Dir(mSetNmtq, mDir, mPfx, mSfx)
End Select
jj.Shw_Dbg cSub, cMod, , "mRslt,mSetNmtq,mDir,mSfx", mRslt, mSetNmtq, mDir, mSfx
If jj.Opn_Dir(mDir) Then Stop
Exit Function
E: Exp_SetNmtq2Dir_Tst = True
End Function
Function Exp_Prj( _
    Optional pLikNmPrj$ = "*" _
    , Optional pAcs As Access.Application = Nothing _
    , Optional pDir$ = "" _
    , Optional pNoMsg As Boolean = False _
    , Optional oErr% _
    , Optional oOK% _
    ) As Boolean
Const cSub$ = "Exp_Prj"
Dim mAnPrj$(): If jj.Fnd_AnPrj(mAnPrj, pLikNmPrj, , pAcs) Then ss.A 1: GoTo E
Dim mErr%, mOK%
Dim mDir$: If pDir = "" Then mDir = jj.Sdir_ExpPgm Else mDir = pDir
If jj.Dlt_Dir(mDir, "*.bas") Then ss.A 2: GoTo E
If jj.Dlt_Dir(mDir, "*.cls") Then ss.A 3: GoTo E
Dim iPrj%
For iPrj = 0 To jj.Siz_Ay(mAnPrj) - 1
    Dim mNmPrj$: mNmPrj = mAnPrj(iPrj)
    Dim mFfnRf$: mFfnRf = mDir & mNmPrj & ".Reference.txt"
    If jj.Exp_Rf(mFfnRf, mNmPrj, pAcs) Then ss.A 4: GoTo E
    If jj.Exp_Md(, mNmPrj, pAcs, mDir, True, True, mErr, mOK) Then ss.A 5: GoTo E
    oErr = oErr + mErr
    oOK = oOK + mOK
    Debug.Print jj.Fmt_Str(mNmPrj & " has exported modules: OK[{0}]  Err[{1}]", mOK, mErr)
Next
If Not pNoMsg Then
    MsgBox jj.Fmt_Str( _
        "[{0}] projects" & vbLf & _
        "[{1}] modules exported OK" & vbLf & _
        "[{2}] modules exported with errors", _
            jj.Siz_Ay(mAnPrj), oOK, oErr), vbInformation, "Modules Export Result"
    jj.Opn_Dir mDir$
End If
Exit Function
E: Exp_Prj = True: ss.B cSub, cMod, "pLikNmPrj, pAcs,pDir$,pNoMsg", pLikNmPrj, jj.ToStr_Acs(pAcs), pDir$, pNoMsg
End Function
Function Exp_Md( _
    Optional pLikNmm$ = "*" _
    , Optional pNmPrj$ = "" _
    , Optional pAcs As Access.Application = Nothing _
    , Optional pDir$ = "" _
    , Optional pNoMsg As Boolean = False _
    , Optional pKeepDir As Boolean = False _
    , Optional oErr% _
    , Optional oOK% _
    ) As Boolean
Const cSub$ = "Exp_Md"
'Aim: Export currentMdb all modules to text files in folder {SmDir} after sorting by procedure/Function Exp_name
'==Start
'Create / Kill all *.txt in .\Mod\
Dim mDir$: If pDir = "" Then mDir = jj.Sdir_ExpPgm Else mDir = pDir
If Not pKeepDir Then
    If jj.Dlt_Dir(mDir, "*.bas") Then ss.A 1: GoTo E
    If jj.Dlt_Dir(mDir, "*.cls") Then ss.A 2: GoTo E
End If
Dim mAcs As Access.Application: Set mAcs = jj.Cv_Acs(pAcs)
Dim mPrj As VBProject: If jj.Cv_Prj(mPrj, pNmPrj, mAcs) Then ss.A 2: GoTo E
Dim mAnm$(): If jj.Fnd_Anm_ByPrj(mAnm, mPrj, pLikNmm, True) Then ss.A 3: GoTo E
''Loop the Collection and call <ExpMd_For1Md>
Dim J%
For J = 0 To jj.Siz_Ay(mAnm) - 1
    Dim mMd As CodeModule: If jj.Fnd_Md(mMd, mPrj, mAnm(J)) Then ss.A 4: GoTo E
    If jj.Exp_Md2Dir(mMd, mDir) Then
        oErr = oErr + 1
    Else
        oOK = oOK + 1
    End If
Next
If Not pNoMsg Then
    MsgBox jj.Fmt_Str( _
        "[{0}] modules exported OK" & vbLf & _
        "[{1}] modules exported with errors", _
            oOK, oErr), vbInformation, "Modules Export Result"
    jj.Opn_Dir mDir
End If
Exit Function
R: ss.R
E: Exp_Md = True: ss.B cSub, cMod, "pLikNmm,pNmPrj,pAcs,pDir,pNoMsg,pKeepDir", pLikNmm, pNmPrj, jj.ToStr_Acs(pAcs), pDir, pNoMsg, pKeepDir
End Function
#If Tst Then
Function Exp_Md_Tst() As Boolean
jj.Exp_Md "Acpt*"
End Function
#End If
Function Exp_Md2Dir(pMd As CodeModule, pDir$) As Boolean
Const cSub$ = "Exp_Md2Dir"
'==Start
'Open the given module {pNmm} in <mCdMdCur>
If TypeName(pMd) = "Nothing" Then ss.A 1, "pMd cannot be nothing": GoTo E
Dim mExt$
Select Case pMd.Parent.Type
Case VBIDE.vbext_ComponentType.vbext_ct_ClassModule:    mExt = ".cls"
Case VBIDE.vbext_ComponentType.vbext_ct_StdModule:      mExt = ".bas"
Case VBIDE.vbext_ComponentType.vbext_ct_Document:       mExt = ".cls"
Case Else: ss.A 2, "Unexpect pMd Type", , "pMd.Typ", jj.ToStr_TypCmp(pMd.Parent.Type): GoTo E
End Select

Dim Fno As Byte: If jj.Opn_Fil_ForOutput(Fno, pDir & jj.ToStr_Md(pMd) & mExt) Then ss.A 2: GoTo E
'Export the non-proceudure lines
If pMd.CountOfDeclarationLines > 0 Then Print #Fno, pMd.Lines(1, pMd.CountOfDeclarationLines)
Dim mAnPrc$(): If jj.Fnd_AnPrc_ByMd(mAnPrc, pMd, , True, True, True) Then ss.A 3: GoTo E
Dim J%
For J = 0 To jj.Siz_Ay(mAnPrc) - 1
    Dim iNmPrc$, iPrcLinBeg$, iPrcLinEnd$, iPrcNLin$
    If jj.Brk_Str_To3Seg(iNmPrc, iPrcLinBeg, iPrcLinEnd, mAnPrc(J)) Then ss.A 4: GoTo E
    Dim mNLin&
    Dim mLinBeg&: mLinBeg = iPrcLinBeg
    Dim mLinEnd&: mLinEnd = iPrcLinEnd
    mNLin = mLinEnd - mLinBeg + 1
    Print #Fno, pMd.Lines(mLinBeg, mNLin)
Next
Close #Fno
Exit Function
R: ss.R
E: Exp_Md2Dir = True: ss.B cSub, cMod, "pMd,pDir", jj.ToStr_Md(pMd), pDir
End Function
Function Exp_Tbl_ToMdb_ByTblOup(pFbTar$) As Boolean
'Aim: Currentdb db's {pLikNmt} tables to {pFbTar}
Const cSub$ = "Exp_Tbl_ToMdb_ByTblOup"
If pFbTar = "" Then ss.A 1: GoTo E
If Not jj.IsTbl("tblOup") Then Exit Function
If VBA.Dir(pFbTar) = "" Then If jj.Crt_Fb(pFbTar) Then ss.A 1: GoTo E
With CurrentDb.TableDefs("tblOput").OpenRecordset
    While Not .EOF
        If InStr(!LikNmt, "*") > 0 Then
            If jj.Snd_Tbl_ToMdb(!LikNmt, pFbTar$) Then ss.A 2: GoTo E
        Else
            Dim mSql$: mSql = jj.Fmt_Str("Select * into {0} in '{1}' from {0}", !LikNmt, pFbTar)
            If jj.Run_Sql(mSql) Then ss.A 5: GoTo E
        End If
        .MoveNext
    Wend
    .Close
End With
Exit Function
R: ss.R
E: Exp_Tbl_ToMdb_ByTblOup = True: ss.B cSub, cMod, "pFbTar", pFbTar
End Function
#If Tst Then
Function Exp_Tbl_ToMdb_ByTblOup_Tst() As Boolean
If jj.Exp_Tbl_ToMdb_ByTblOup("C:\aa.mdb") Then Stop
End Function
#End If

