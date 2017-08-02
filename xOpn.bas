Attribute VB_Name = "xOpn"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xOpn"
Function Opn_Wb_ByDirLsFn(pDir$, pLsFn$) As Boolean
Excel.Application.Workbooks.Close
Dim mAyFn$(): mAyFn = Split(pLsFn, ",")
Dim J%
For J = 0 To jj.Siz_Ay(mAyFn) - 1
    Dim mFfn$: mFfn = pDir & mAyFn(J) & ".xls"
    If jj.IsFfn(mFfn) Then Set_Wb_Min Excel.Application.Workbooks.Open(mFfn)
Next
Set_WbAll_Min
End Function
Function Opn_Calendar(Optional pFy As Byte = 0) As Boolean
Const cSub$ = "Opn_Calendar"
Dim mFY As Byte: mFY = IIf(pFy = 0, jj.Cv_Dte2FyNo, pFy)
Dim mFfn$: mFfn = jj.Sdir_Wrk & "Calendar\SPL Company Calendar FY20" & Format(mFY, "00") & ".xls"
Dim mWb As Workbook
Opn_Calendar = jj.Opn_Wb_R(mWb, mFfn, True, True)
End Function
Function Opn_CurDb(pAcs As Access.Application, pFb$, Optional pIsExcl As Boolean = False, Optional pPwd$ = "", Optional pVisible As Boolean = False) As Boolean
Const cSub$ = "Opn_CurDb"
On Error GoTo Nxt
If pAcs.CurrentDb.Name = pFb Then Exit Function
Nxt:
jj.Cls_CurDb pAcs
On Error GoTo R
With pAcs
    .OpenCurrentDatabase pFb, pIsExcl, pPwd
    If .Visible <> pVisible Then .Visible = pVisible
End With
Exit Function
R: ss.R
E: Opn_CurDb = True: ss.B cSub, cMod, "pFb,pIsExcl,Any Pwd", pFb, pIsExcl, pPwd <> ""
End Function
#If Tst Then
Function Opn_CurDb_Tst() As Boolean
Const cFb$ = "P:\WorkingDir\PgmObj\Lgc\LgcExpDb.mdb"
Dim mAcs As Access.Application: Set mAcs = g.gAcs
Debug.Print Time
Dim J%
For J = 1 To 20
    If jj.Opn_CurDb(mAcs, cFb) Then Stop
    'If jj.Cls_CurDb(mAcs) Then Stop
    Debug.Print J, Time
Next
mAcs.Visible = True
End Function
#End If
Function Opn_Db_Txt(oDb As DAO.Database, pDir$) As Boolean
'Aim: Open {pDir} as a database by referring all *.txt as table
'Note: Schema.ini in {pDir} will be used if exist.  See jj.Cv_Fdf2Schema() about Schema.ini
Const cSub$ = "Opn_Db_Txt"
On Error GoTo R
Set oDb = g.gDbEng.OpenDatabase(pDir, False, False, "Text;Database=" & pDir)
Exit Function
R: ss.R
E: Opn_Db_Txt = True: ss.B cSub, cMod, "pDir", pDir$
End Function
#If Tst Then
Function Opn_Db_Txt_Tst() As Boolean
'Const cDir$ = "X:\Ldb\WorkingDir\Data"
'Const cDir$ = "D:\Data\Johnson Cheung\MyDoc\My Projects\My Projects Library\Ldb\Ldb\WorkingDir\Data"
Const cDir$ = "c:\"
Dim mDb As DAO.Database: If jj.Opn_Db_Txt(mDb, cDir) Then Stop
Debug.Print mDb.TableDefs.Count
'CurrentDb.Execute "Select * into [XXX] from A1#txt in '' [Text;DataBase=" & cDir & "]"
DoCmd.TransferText acExportHTML, , "a1", "c:\aa#html", True
'CurrentDb.Execute "Select * into [tblIIM] from IIM#txt in '' [Text;DataBase=" & cDir & "]"
mDb.Close
End Function
#End If
Function Opn_Db(oDb As DAO.Database, pFb$, pIsReadOnly As Boolean) As Boolean
Const cSub$ = "Opn_Db"
On Error GoTo R
If jj.Cls_Db(oDb) Then ss.A 1: GoTo E
Set oDb = g.gDbEng.OpenDatabase(pFb, , pIsReadOnly)
Exit Function
R: ss.R
E: Opn_Db = True: ss.B cSub, cMod, "pFb,pIsReadOnly", pFb, pIsReadOnly
End Function
Function Opn_Db_R(oDb As DAO.Database, pFb$) As Boolean
If pFb = "" Then Set oDb = CurrentDb: Exit Function
Opn_Db_R = Opn_Db(oDb, pFb, True)
End Function
Function Opn_Db_RW(oDb As DAO.Database, pFb$) As Boolean
If pFb = "" Then Set oDb = CurrentDb: Exit Function
Opn_Db_RW = Opn_Db(oDb, pFb, False)
End Function
Function Opn_Dir(pDir$) As Boolean
Const cSub$ = "Opn_Dir"
If Not jj.IsDir(pDir) Then ss.A 1: GoTo E
Shell jj.Fmt_Str("Explorer.exe ""{0}""", pDir), vbMaximizedFocus
Exit Function
R: ss.R
E: Opn_Dir = True: ss.B cSub, cMod, ""
End Function
Function Opn_DirCur() As Boolean
Opn_DirCur = jj.Opn_Dir(jj.Sdir_Hom)
End Function
Function Opn_DirRpt() As Boolean
Opn_DirRpt = jj.Opn_Dir(jj.Sdir_Rpt)
End Function
Function Opn_Fil_ForInput(oFno As Byte, pFfn$) As Boolean
Const cSub$ = "Opn_Fil_ForInput"
On Error GoTo R
oFno = FreeFile: Open pFfn For Input As #oFno
Exit Function
R: ss.R
E: Opn_Fil_ForInput = True: ss.B cSub, cMod, "pFfn", pFfn
End Function
Function Opn_Fil_ForOutput(oFno As Byte, pFfn$, Optional pOvrWrt As Boolean = False) As Boolean
Const cSub$ = "Opn_Fil_ForOutput"
If jj.Ovr_Wrt(pFfn, pOvrWrt) Then ss.A 1: GoTo E
On Error GoTo R
oFno = FreeFile: Open pFfn For Output As #oFno
Exit Function
R: ss.R
E: Opn_Fil_ForOutput = True: ss.B cSub, cMod, "pFfn", pFfn
End Function
#If Tst Then
Function Opn_Fil_ForOutput_Tst() As Boolean
Const cFt$ = "c:\aa.csv"
Dim mF As Byte: If jj.Opn_Fil_ForOutput(mF, cFt, True) Then Stop: GoTo E
Print #mF, "aa,bb"
Close #mF
Dim mWb As Workbook: If jj.Opn_Wb_R(mWb, cFt, , True) Then Stop: GoTo E
Exit Function
E: Opn_Fil_ForOutput_Tst = True
End Function
#End If
Function Opn_Frm(pNmFrm$, Optional pOpnArgs, Optional pIsDialog As Boolean = False, Optional oFrm As Access.Form) As Boolean
Const cSub$ = "Opn_Frm"
On Error GoTo R
If pIsDialog Then
    DoCmd.OpenForm pNmFrm, , , , , acDialog, pOpnArgs
    Exit Function
End If
DoCmd.OpenForm pNmFrm, , , , , , pOpnArgs
Set oFrm = Access.Application.Forms(pNmFrm)
Exit Function
R: ss.R
E: Opn_Frm = True: ss.B cSub, cMod, "pNmFrm", pNmFrm
End Function
#If Tst Then
Function Opn_Frm_Tst() As Boolean
Const cNmFrm$ = "frmIIC_Tst"
If jj.Opn_Frm(cNmFrm) Then Stop: GoTo E
Stop
If jj.Cls_Frm(cNmFrm) Then Stop: GoTo E
Exit Function
E: Opn_Frm_Tst = True
End Function
#End If
Function Opn_PDF(pFfnPDF$) As Boolean
Shell jj.Fmt_Str("""C:\Program Files\Adobe\Acrobat 7.0\Reader\AcroRd32.exe"" ""{0}""", pFfnPDF), vbMaximizedFocus
End Function
Function Opn_Qry(pNmq$, Optional pView As AcView = AcView.acViewNormal, Optional pOpnDtaMd As AcOpenDataMode = AcOpenDataMode.acEdit) As Boolean
Const cSub$ = "Opn_Qry"
On Error GoTo R
DoCmd.OpenQuery pNmq, pView, pOpnDtaMd
DoCmd.SelectObject acQuery, pNmq
Exit Function
R: ss.R
E: Opn_Qry = True: ss.B cSub, cMod, "pNmq,pView,pOpnDtaMd", pNmq, pView, pOpnDtaMd
End Function
Function Opn_ReadMe(Optional pNmRpt$ = "") As Boolean
Const cSub$ = "Opn_ReadMe"
Dim mNmRpt$: If pNmRpt <> "" Then mNmRpt = pNmRpt & "_"
Dim mFfnReadMe$: mFfnReadMe = jj.Sdir_Wrk & mNmRpt & "ReadMe.xls"
If VBA.Dir(mFfnReadMe) <> "" Then
    Dim mWb As Workbook: If jj.Opn_Wb_R(mWb, mFfnReadMe, True) Then ss.A 1: GoTo E
    mWb.Application.Visible = True
    Exit Function
End If

mFfnReadMe = jj.Sdir_Wrk & mNmRpt & "ReadMe.doc"
If VBA.Dir(mFfnReadMe) <> "" Then
    Dim mWrd As Word.Document: If jj.Opn_Wrd_R(mWrd, mFfnReadMe, True) Then ss.A 2: GoTo E
    mWrd.Application.Visible = True
    Exit Function
End If
ss.A 1, "No ReadMe file is found": GoTo E
R: ss.R
E: Opn_ReadMe = True: ss.B cSub, cMod, "Dir,Report Name,Read Me file name", jj.Sdir_Wrk, pNmRpt, mFfnReadMe
End Function
Function Opn_Rpt(pNmRptSht$, Optional pNmSess$ = "", Optional pTimStampOpt As eTimStampOpt = eDte, Optional pInRptDir As Boolean = False, Optional pExt$ = ".xls") As Boolean
Const cSub$ = "Opn_Rpt"
Dim mFfnRpt$: mFfnRpt = jj.Sffn_Rpt(pNmRptSht, pNmSess, pTimStampOpt, pInRptDir, pExt)
If Not jj.IsFfn(mFfnRpt) Then ss.A 1: GoTo E
Select Case pExt
Case ".xls"
    Dim mWb As Workbook: If jj.Opn_Wb_RW(mWb, mFfnRpt, True, True) Then ss.A 1: GoTo E
    Exit Function
Case ".doc"
    Dim mWrd As Word.Document: If jj.Opn_Wrd_RW(mWrd, mFfnRpt, True, True) Then ss.A 2: GoTo E
    Exit Function
Case ".pdf"
    If jj.Opn_PDF(mFfnRpt) Then ss.A 3: GoTo E
Case Else
    ss.A 4, "Extension of the report must be .xls, .doc or .pdf": GoTo E
End Select
Exit Function
E: Opn_Rpt = True: ss.B cSub, cMod, "pNmRptSht,pNmSess,pTimStampOpt,pInRptDir,pExt", pNmRptSht, pNmSess, pTimStampOpt, pInRptDir, pExt
End Function
Function Opn_Rs(oRs As DAO.Recordset, pSql$, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Opn_Rs"
On Error GoTo R
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Set oRs = mDb.OpenRecordset(pSql)
Exit Function
R: ss.R
E: Opn_Rs = True: ss.B cSub, cMod, "pSql", pSql
End Function
Function Opn_Rs_ByNmq(oRs As DAO.Recordset, pNmq$) As Boolean
Const cSub$ = "Opn_Rs_ByNmq"
On Error GoTo R
Set oRs = CurrentDb.QueryDefs(pNmq).OpenRecordset
Exit Function
R: ss.R
E: Opn_Rs_ByNmq = True: ss.B cSub, cMod, "pNmq", pNmq
End Function
Function Opn_Sql(pSql$, Optional pNmq$ = "qry") As Boolean
'Aim: Create or set sql of given qQryNam and open to preview.  Usually for debug
jj.Crt_Qry pNmq, pSql
DoCmd.OpenQuery pNmq, acViewPreview, acReadOnly
End Function
Function Opn_Tbl(pNmt$, Optional pRw As Boolean = False) As Boolean
Const cSub$ = "Opn_Tbl"
On Error GoTo R
If pRw Then
    DoCmd.OpenTable pNmt, , acReadOnly
Else
    DoCmd.OpenTable pNmt
End If
DoCmd.SelectObject acTable, pNmt
Exit Function
R: ss.R
E: Opn_Tbl = True: ss.B cSub, cMod, "pNmt,pRw", pNmt, pRw
End Function
Function Opn_SetNmt(pSetNmt$, Optional pRw As Boolean = False) As Boolean
Const cSub$ = "Opn_SetNmt"
On Error GoTo R
Dim mAnt$()
If jj.Fnd_Ant_BySetNmt(mAnt, pSetNmt) Then ss.A 1: GoTo E
Dim J%, N%: N = jj.Siz_Ay(mAnt)
If pRw Then
    For J = 0 To N - 1
        DoCmd.OpenTable mAnt(J), , acReadOnly
    Next
Else
    For J = 0 To N - 1
        DoCmd.OpenTable mAnt(J)
    Next
End If
If N > 0 Then DoCmd.SelectObject acTable, mAnt(0)
Exit Function
R: ss.R
E: Opn_SetNmt = True: ss.B cSub, cMod, "pSetNmt,pRw", pSetNmt, pRw
End Function
Function Opn_Tbl_ByPfx(pPfx$) As Boolean
Dim L%: L = Len(pPfx)
Dim iTbl As TableDef: For Each iTbl In CurrentDb.TableDefs
    If Left(iTbl.Name, L) = pPfx Then DoCmd.OpenTable iTbl.Name, , acReadOnly
Next
Stop
For Each iTbl In CurrentDb.TableDefs
    If Left(iTbl.Name, L) = pPfx Then Call DoCmd.Close(acTable, iTbl.Name, acSaveNo)
Next
End Function
Function Opn_Ws(oWb As Workbook, oWs As Worksheet, pFx$, pNmWs$, Optional pIsReadOnly As Boolean = False, Optional pIsInNewXls As Boolean = False, Optional pIsVisible As Boolean = False) As Boolean
Const cSub$ = "Opn_Ws"
On Error GoTo R
If jj.Opn_Wb(oWb, pFx, pIsReadOnly, pIsInNewXls, pIsVisible) Then ss.A 1: GoTo E
If Not jj.IsWs(oWb, pNmWs) Then oWb.Close: ss.A 2: GoTo E
Set oWs = oWb.Sheets(pNmWs)
Exit Function
R: ss.R
E: Opn_Ws = True: ss.B cSub, cMod, "pFx,pNmWs,pIsReadOnly,pIsInNewXls", pFx, pNmWs, pIsReadOnly, pIsInNewXls
End Function
Function Opn_Ws_Tst() As Boolean
Dim mWb As Workbook, mWs As Worksheet
Dim mFx$: mFx = "c:\tmp\aa.xls"
Dim mNmWs$: mNmWs = "xxx"
If jj.Crt_Wb(mWb, mFx, True, mNmWs) Then Stop
If jj.Cls_Wb(mWb, True) Then Stop
If jj.Opn_Ws_R(mWb, mWs, mFx, mNmWs) Then Stop
Stop
End Function
Function Opn_Wb(oWb As Workbook, pFx$, Optional pIsReadOnly As Boolean = False, Optional pIsInNewXls As Boolean = False, Optional pIsVisible As Boolean = False) As Boolean
Const cSub$ = "Opn_Wb"
On Error GoTo R
Dim mXls As Excel.Application
If pIsInNewXls Then
    Set mXls = New Excel.Application
Else
    Set mXls = jj.g.gXls
End If
Set oWb = mXls.Workbooks.Open(pFx, , pIsReadOnly)
If pIsVisible Then mXls.Visible = True
Exit Function
R: ss.R
E: Opn_Wb = True: ss.B cSub, cMod, "pFx,pIsReadOnly,pIsInNewXls", pFx, pIsReadOnly, pIsInNewXls
End Function
Function Opn_Wb_ByWbs(pWbs As Excel.Workbooks, pFx$, pIsReadOnly As Boolean) As Boolean
Const cSub$ = "Opn_Wb_ByWbs"
On Error GoTo R
pWbs.Open pFx, , pIsReadOnly
Exit Function
R: ss.R
E: Opn_Wb_ByWbs = True: ss.B cSub, cMod, "pFx,pIsReadOnly", pFx, pIsReadOnly
End Function
Function Opn_Wb_ByWbs_R(pWbs As Excel.Workbooks, pFx$) As Boolean
Opn_Wb_ByWbs_R = jj.Opn_Wb_ByWbs(pWbs, pFx, True)
End Function
Function Opn_Wb_ByWbs_RW(pWbs As Excel.Workbooks, pFx$) As Boolean
Opn_Wb_ByWbs_RW = jj.Opn_Wb_ByWbs(pWbs, pFx, False)
End Function
Function Opn_Wb_R(oWb As Workbook, pFx$, Optional pIsInNewXls As Boolean = False, Optional pIsVisible As Boolean = False) As Boolean
Opn_Wb_R = jj.Opn_Wb(oWb, pFx, True, pIsInNewXls, pIsVisible)
End Function
Function Opn_Wb_RW(oWb As Workbook, pFx$, Optional pIsInNewXls As Boolean = False, Optional pIsVisible As Boolean = False) As Boolean
Opn_Wb_RW = jj.Opn_Wb(oWb, pFx, False, pIsInNewXls, pIsVisible)
End Function
Function Opn_Ws_R(oWb As Workbook, oWs As Worksheet, pFx$, pNmWs$, Optional pIsInNewXls As Boolean = False, Optional pIsVisible As Boolean = False) As Boolean
Opn_Ws_R = jj.Opn_Ws(oWb, oWs, pFx, pNmWs, True, pIsInNewXls, pIsVisible)
End Function
Function Opn_Ws_RW(oWb As Workbook, oWs As Worksheet, pFx$, pNmWs$, Optional pIsInNewXls As Boolean = False, Optional pIsVisible As Boolean = False) As Boolean
Opn_Ws_RW = jj.Opn_Ws(oWb, oWs, pFx, pNmWs, False, pIsInNewXls, pIsVisible)
End Function
Function Opn_Wrd(oWrd As Word.Document, pFfnWrd$, pIsReadOnly As Boolean, pIsInNewWrd As Boolean, Optional pIsVisible As Boolean = False) As Boolean
Const cSub$ = "Opn_Wrd"
On Error GoTo R
Dim mWrd As Word.Application
If pIsInNewWrd Then
    Set mWrd = New Word.Application
Else
    Set mWrd = gWrd
End If
Set oWrd = mWrd.Documents.Open(pFfnWrd, , pIsReadOnly)
If pIsVisible Then mWrd.Visible = True
Exit Function
R: ss.R
E: Opn_Wrd = True: ss.B cSub, cMod, "pFfnWrd,pIsReadOnly,pIsInNewWrd", pFfnWrd, pIsReadOnly, pIsInNewWrd
End Function
Function Opn_Wrd_R(oWrd As Word.Document, pFfnWrd$, Optional pIsInNewWrd As Boolean = False, Optional pIsVisible As Boolean = False) As Boolean
Opn_Wrd_R = jj.Opn_Wrd(oWrd, pFfnWrd, True, pIsInNewWrd, pIsVisible)
End Function
Function Opn_Wrd_RW(oWrd As Word.Document, pFfnWrd$, Optional pIsInNewWrd As Boolean = False, Optional pIsVisible As Boolean = False) As Boolean
Opn_Wrd_RW = jj.Opn_Wrd(oWrd, pFfnWrd, False, pIsInNewWrd, pIsVisible)
End Function
Function Opn_Xls_InDir(pDir$, Optional pFspc$ = "*.xls") As Boolean
'Aim: Open all Excel files in {pDir} with given {pFspc}
Const cSub$ = "Opn_Xls_InDir"
On Error GoTo R
Dim mAyFn$(): If jj.Fnd_AyFn(mAyFn, pDir, pFspc) Then ss.A 1: GoTo E
Dim mA$, mXls As New Excel.Application
Dim mWbs As Workbooks: Set mWbs = mXls.Workbooks
Dim J%: For J = 0 To jj.Siz_Ay(mAyFn) - 1
    If jj.Opn_Wb_ByWbs_R(mWbs, pDir & mAyFn(J)) Then mA = jj.Add_Str(mA, mAyFn(J))
Next
If mA <> "" Then ss.A 1, "Some Excel files cannot be opened", , "Dir,Files cannot be openned", pDir, mA
mXls.Windows.Arrange xlArrangeStyleTiled
mXls.Visible = True
Exit Function
R: ss.R
E: Opn_Xls_InDir = True: ss.B cSub, cMod, ""
End Function
Public Function Opn_Fx(pFx$) As Boolean
On Error GoTo R
If Dir(pFx) = "" Then MsgBox "File not found" & vbLf & pFx: GoTo E
Dim mXls As New Excel.Application
mXls.Workbooks.Open pFx
mXls.Visible = True
Exit Function
R: ss.R
E: Opn_Fx = True
End Function
Function Opn_Fx_FmTpWithRfh(ByRef oWb As Workbook, pFx$, pFxTp$, Optional pStopVisible As Boolean = False) As Boolean
'Aim: Return true for error.
'     Cpy {pFxTp} to {pFx}
'     Refresh data in {pFx} from CurrentDb
'     Opn pFx in {oWb}
'     Optionally set oWb.application visible depending on {pStopVisible}
'     Prompt error
On Error GoTo X
If xCpy.Cpy_Fil(pFxTp, pFx, pOvrWrt:=True) Then GoTo E
Dim mXls As New Excel.Application
Set oWb = mXls.Workbooks.Open(pFx)
xRfh.Rfh_Wb oWb
If Not pStopVisible Then mXls.Visible = True
Exit Function
X: MsgBox Err.Description
E: Opn_Fx_FmTpWithRfh = True
    On Error Resume Next
    mXls.DisplayAlerts = False
    oWb.Close False
    mXls.Quit
    Set mXls = Nothing
End Function
Public Function Opn_Fw(pFw$) As Boolean
On Error GoTo R
If Dir(pFw) = "" Then MsgBox "File not found" & vbLf & pFw: GoTo E
Dim mWrd As New Word.Application
mWrd.Documents.Open pFw
mWrd.Visible = True
Exit Function
R: ss.R
E: Opn_Fw = True
End Function
