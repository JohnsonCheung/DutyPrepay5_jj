Attribute VB_Name = "xRun"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xRun"
Function Run_Sql_By_Repeat_ByAm(pSqlTp$, pAm() As tMap, Optional pNmqPfx$ = "", Optional pAcs As Access.Application = Nothing) As Boolean
Const cSub$ = "Run_Sql_By_Repeat_ByLm"
On Error GoTo R
Dim mAcs As Access.Application, mDb As DAO.Database
Set mAcs = jj.Cv_Acs(pAcs): Set mDb = mAcs.CurrentDb
Dim mAySql$(): If jj.Fmt_Str_Repeat_ByAm_IntoAy(mAySql, pSqlTp, pAm) Then ss.A 1: GoTo E
Dim J%
Dim N%: N = jj.Siz_Ay(mAySql)
If pNmqPfx <> "" Then
    For J = 0 To N - 1
        If jj.Crt_Qry(pNmqPfx & Format(J, "00"), mAySql(J), mDb) Then ss.A 1: GoTo E
    Next
End If
For J = 0 To N - 1
    If jj.Run_Sql(mAySql(J), mAcs) Then ss.A 2: GoTo E
Next
Exit Function
R: ss.R
E: Run_Sql_By_Repeat_ByAm = True: ss.B cSub, cMod, "pNmqPfx,pSqlTp,pAm", pNmqPfx, pSqlTp, jj.ToStr_Am(pAm)
X: Set mDb = Nothing
End Function
Function Run_Sql_By_Repeat_ByLm(pSqlTp$, pLm$, Optional pNmqPfx$ = "", Optional pAcs As Access.Application = Nothing) As Boolean
Run_Sql_By_Repeat_ByLm = Run_Sql_By_Repeat_ByAm(pSqlTp, Get_Am_ByLm(pLm), pNmqPfx, pAcs)
End Function
Function Run_Lgs(oTrc&, pNmLgs$, Optional pLm$) As Boolean
Const cSub$ = "Run_Lgs"
If jj.Crt_LgsTrc(oTrc, pNmLgs, pLm) Then ss.A 1: GoTo E
If jj.Run_Lgs_ByTrc(oTrc) Then ss.A 2: GoTo E
Exit Function
R: ss.R
E: Run_Lgs = True: ss.B cSub, cMod, "pNmLgs,pLm", pNmLgs, pLm
End Function
#If Tst Then
Function Run_Lgs_Tst() As Boolean
Dim mTrc&
If jj.Crt_Tbl_FmLnkLnt("D:\SPLHalfWayHouse\MPSDetail\VerNew@2007_01_04\MPS.mdb", "tblLgs,tblLgsLgc,tblHst_Lgs,tblHst_LgsLgc") Then Stop
If jj.Run_Lgs(mTrc, "MPS", "Env=FE,Brand=TH") Then Stop Else Debug.Print mTrc
If jj.Run_Lgs(mTrc, "MPS", "Env=FEPROD,Brand=Coach") Then Stop Else Debug.Print mTrc
If jj.Run_Lgs(mTrc, "MPS", "Env=FEPROD,Brand=ESQ") Then Stop Else Debug.Print mTrc
If jj.Run_Lgs(mTrc, "MPS", "Env=FEPROD,Brand=Juicy") Then Stop Else Debug.Print mTrc
If jj.Run_Lgs(mTrc, "MPS", "Env=CHPROD,Brand=HB") Then Stop Else Debug.Print mTrc
If jj.Run_Lgs(mTrc, "MPS", "Env=CHPROD,Brand=Lacoste") Then Stop Else Debug.Print mTrc
If jj.Run_Lgs(mTrc, "MPS", "Env=NAPROD,Brand=TH") Then Stop Else Debug.Print mTrc
If jj.Run_Lgs(mTrc, "MPS", "Env=NAPROD,Brand=Coach") Then Stop Else Debug.Print mTrc
If jj.Run_Lgs(mTrc, "MPS", "Env=NAPROD,Brand=ESQ") Then Stop Else Debug.Print mTrc
If jj.Run_Lgs(mTrc, "MPS", "Env=NAPROD,Brand=Juicy") Then Stop Else Debug.Print mTrc
If jj.Run_Lgs(mTrc, "MPS", "Env=NAPROD,Brand=Lacoste") Then Stop Else Debug.Print mTrc
If jj.Run_Lgs(mTrc, "MPS", "Env=NAPROD,Brand=HB") Then Stop Else Debug.Print mTrc
End Function
#End If
Function Run_Lgs_ByTrc(pTrc&) As Boolean
'Aim: Pick up AnLgc(), AnLp(), AnLv() From [tblHst_Lgs] & [tblHst_LgsLgc] by {pTrc} and run them
Const cSub$ = "Run_Lgs_ByTrc"
'Do Pick up AnLgc(), AnLp(), AnLv() From [tblHst_Lgs] & [tblHst_LgsLgc] by {pTrc}
Dim mAnLgc$(), mAyLm$()
Do
    Dim mSql$: mSql = "SELECT NmLgc, Lp, Lv" & _
        " FROM tblHst_LgsLgc" & _
        " Where Trc=" & pTrc & _
        " And   Trc in (Select Trc from tblHst_Lgs where IsNull(DteBeg) or IsNull(DteEnd))" & _
        " And   (IsNull(DteBeg) or IsNull(DteEnd))" & _
        " Order by Sno"
    If jj.Fnd_LoAyV_FmSql(mSql, "NmLgc,Lm", mAnLgc, mAyLm) Then ss.A 1: GoTo E
Loop Until True

Do
    mSql = "Update tblHst_Lgs set DteBeg=Now where Trc=" & pTrc
    If jj.Run_Sql(mSql) Then ss.A 2: GoTo E
    Dim J%
    For J = 0 To jj.Siz_Ay(mAnLgc) - 1
        mSql = jj.Fmt_Str("Update tblHst_LgsLgc set DteBeg=Now where NmLgc='{0}' and Trc={1}", mAnLgc(J), pTrc)
        If jj.Run_Sql(mSql) Then ss.A 3: GoTo E

        If jj.Run_Lgc(mAnLgc(J), pTrc, mAyLm(J)) Then ss.A 4, "Error in one of the steps of a Lgs", , "NmLgc,J with Err,Lm", mAnLgc(J), J, mAyLm(J): GoTo E
        
        mSql = jj.Fmt_Str("Update tblHst_LgsLgc set DteEnd=Now where NmLgc='{0}' and Trc={1}", mAnLgc(J), pTrc)
        If jj.Run_Sql(mSql) Then ss.A 5: GoTo E
    Next
    mSql = "Update tblHst_Lgs set DteEnd=Now where Trc=" & pTrc
    If jj.Run_Sql(mSql) Then ss.A 6: GoTo E
Loop Until True
Exit Function
R: ss.R
E: Run_Lgs_ByTrc = True: ss.B cSub, cMod, "pTrc", pTrc
End Function
#If Tst Then
Function Run_Lgs_ByTrc_Tst() As Boolean
Dim mCase As Byte: mCase = 1
Select Case mCase
Case 1
    If jj.Run_Lgs_ByTrc(12) Then Stop
Case 2
    g.gIsBch = True
    Dim oTrc&
    For oTrc = 18 To 29
        jj.Run_Lgs_ByTrc oTrc
    Next
    g.gIsBch = False
End Select
End Function
#End If
Function Run_Lgc(pNmLgc$, Optional pTrc& = 0, Optional pLm$ = "") As Boolean
'Aim: Copy {pNmLgc} in [jj.Sdir_PgmLgc] to [jj.Sffn_TmpLgc] if need
'     Run {pNmLgc} in [jj.Sffn_TmpLgc] with {pTrc,pLn,pLv} as parameters.
Const cSub$ = "Run_Lgc"
' Do Find {mFbOldQsTmp} & 'build' it
Dim mFbOldQsTmp$: If jj.Fnd_Sffn_LgcMdbTmp(mFbOldQsTmp, pNmLgc) Then ss.A 1: GoTo E
Do
    Dim mFbLgc$:    If jj.Fnd_Sffn_LgcMdb(mFbLgc, pNmLgc) Then ss.A 2: GoTo E
    If VBA.Dir(mFbOldQsTmp) = "" Then
        If jj.Cpy_Fil(mFbLgc, mFbOldQsTmp) Then ss.A 4: GoTo E
    Else
        Dim mIsOk As Boolean
        If jj.Chk_LgcVer(mIsOk, mFbOldQsTmp, mFbLgc) Then ss.A 5: GoTo E
        If Not mIsOk Then If jj.Cpy_Fil(mFbLgc, mFbOldQsTmp, True) Then ss.A 6: GoTo E
    End If
Loop Until True
Dim mRslt As Boolean
If jj.Run_Fb(mRslt, mFbOldQsTmp, pNmLgc, pTrc, pLm, Not jj.SysCfg_LgcHidAcs) Then ss.A 7: GoTo E
If mRslt Then ss.A 6, "Return error in Running mFbOldQsTmp", , "mFbOldQsTmp", mFbOldQsTmp: GoTo E
Exit Function
R: ss.R
E: Run_Lgc = True: ss.B cSub, cMod, "pNmLgc,pTrc,pLm", pNmLgc, pTrc, pLm
End Function
#If Tst Then
Function Run_Lgc_Tst() As Boolean
Dim mCase%
For mCase = 31 To 431
    Select Case mCase
    Case 1: If jj.Run_Lgc("ExpTbl") Then Stop
    Case 2: If jj.Run_Lgc("ExpFld") Then Stop
    Case 31: If jj.Run_Lgc("Odbc", 1, "Env=FEPROD,Brand=ESQ") Then Stop
    Case 32: If jj.Run_Lgc("Odbc", 2, "Env=FEPROD,Brand=TH") Then Stop
    Case 33: If jj.Run_Lgc("Odbc", 3, "Env=FEPROD,Brand=HB") Then Stop
    Case 34: If jj.Run_Lgc("Odbc", 4, "Env=FEPROD,Brand=Lacoste") Then Stop
    Case 35: If jj.Run_Lgc("Odbc", 5, "Env=FEPROD,Brand=Coach") Then Stop
    Case 36: If jj.Run_Lgc("Odbc", 6, "Env=FEPROD,Brand=Juicy") Then Stop
    Case 37: If jj.Run_Lgc("Odbc", 7, "Env=NAPROD,Brand=ESQ") Then Stop
    Case 38: If jj.Run_Lgc("Odbc", 8, "Env=NAPROD,Brand=TH") Then Stop
    Case 39: If jj.Run_Lgc("Odbc", 9, "Env=NAPROD,Brand=HB") Then Stop
    Case 40: If jj.Run_Lgc("Odbc", 10, "Env=NAPROD,Brand=Lacoste") Then Stop
    Case 41: If jj.Run_Lgc("Odbc", 11, "Env=NAPROD,Brand=Coach") Then Stop
    Case 42: If jj.Run_Lgc("Odbc", 12, "Env=NAPROD,Brand=Juicy") Then Stop

    Case 131: If jj.Run_Lgc("RfhCusGp", 1, "Brand=ESQ") Then Stop
    Case 132: If jj.Run_Lgc("RfhCusGp", 2, "Brand=TH") Then Stop
    Case 133: If jj.Run_Lgc("RfhCusGp", 3, "Brand=HB") Then Stop
    Case 134: If jj.Run_Lgc("RfhCusGp", 4, "Brand=Lacoste") Then Stop
    Case 135: If jj.Run_Lgc("RfhCusGp", 5, "Brand=Coach") Then Stop
    Case 136: If jj.Run_Lgc("RfhCusGp", 6, "Brand=Juicy") Then Stop
    Case 137: If jj.Run_Lgc("RfhCusGp", 7, "Brand=ESQ") Then Stop
    Case 138: If jj.Run_Lgc("RfhCusGp", 8, "Brand=TH") Then Stop
    Case 139: If jj.Run_Lgc("RfhCusGp", 9, "Brand=HB") Then Stop
    Case 140: If jj.Run_Lgc("RfhCusGp", 10, "Brand=Lacoste") Then Stop
    Case 141: If jj.Run_Lgc("RfhCusGp", 11, "Brand=Coach") Then Stop
    Case 142: If jj.Run_Lgc("RfhCusGp", 12, "Brand=Juicy") Then Stop

    Case 231: If jj.Run_Lgc("RfhCusGp", 1, "Brand=ESQ") Then Stop
    Case 232: If jj.Run_Lgc("RfhCusGp", 2, "Brand=TH") Then Stop
    Case 233: If jj.Run_Lgc("RfhCusGp", 3, "Brand=HB") Then Stop
    Case 234: If jj.Run_Lgc("RfhCusGp", 4, "Brand=Lacoste") Then Stop
    Case 235: If jj.Run_Lgc("RfhCusGp", 5, "Brand=Coach") Then Stop
    Case 236: If jj.Run_Lgc("RfhCusGp", 6, "Brand=Juicy") Then Stop
    Case 237: If jj.Run_Lgc("RfhCusGp", 7, "Brand=ESQ") Then Stop
    Case 238: If jj.Run_Lgc("RfhCusGp", 8, "Brand=TH") Then Stop
    Case 239: If jj.Run_Lgc("RfhCusGp", 9, "Brand=HB") Then Stop
    Case 240: If jj.Run_Lgc("RfhCusGp", 10, "Brand=Lacoste") Then Stop
    Case 241: If jj.Run_Lgc("RfhCusGp", 11, "Brand=Coach") Then Stop
    Case 242: If jj.Run_Lgc("RfhCusGp", 12, "Brand=Juicy") Then Stop
    
    Case 231: If jj.Run_Lgc("RfhFc", 1, "Env=FEPROD,Brand=ESQ") Then Stop
    Case 232: If jj.Run_Lgc("RfhFc", 2, "Env=FEPROD,Brand=TH") Then Stop
    Case 233: If jj.Run_Lgc("RfhFc", 3, "Env=FEPROD,Brand=HB") Then Stop
    Case 234: If jj.Run_Lgc("RfhFc", 4, "Env=FEPROD,Brand=Lacoste") Then Stop
    Case 235: If jj.Run_Lgc("RfhFc", 5, "Env=FEPROD,Brand=Coach") Then Stop
    Case 236: If jj.Run_Lgc("RfhFc", 6, "Env=FEPROD,Brand=Juicy") Then Stop
    Case 237: If jj.Run_Lgc("RfhFc", 7, "Env=NAPROD,Brand=ESQ") Then Stop
    Case 238: If jj.Run_Lgc("RfhFc", 8, "Env=NAPROD,Brand=TH") Then Stop
    Case 239: If jj.Run_Lgc("RfhFc", 9, "Env=NAPROD,Brand=HB") Then Stop
    Case 240: If jj.Run_Lgc("RfhFc", 10, "Env=NAPROD,Brand=Lacoste") Then Stop
    Case 241: If jj.Run_Lgc("RfhFc", 11, "Env=NAPROD,Brand=Coach") Then Stop
    Case 242: If jj.Run_Lgc("RfhFc", 12, "Env=NAPROD,Brand=Juicy") Then Stop
    
    Case 331: If jj.Run_Lgc("GenDta", 1, "Env=FEPROD,Brand=ESQ") Then Stop
    Case 332: If jj.Run_Lgc("GenDta", 2, "Env=FEPROD,Brand=TH") Then Stop
    Case 333: If jj.Run_Lgc("GenDta", 3, "Env=FEPROD,Brand=HB") Then Stop
    Case 334: If jj.Run_Lgc("GenDta", 4, "Env=FEPROD,Brand=Lacoste") Then Stop
    Case 335: If jj.Run_Lgc("GenDta", 5, "Env=FEPROD,Brand=Coach") Then Stop
    Case 336: If jj.Run_Lgc("GenDta", 6, "Env=FEPROD,Brand=Juicy") Then Stop
    Case 337: If jj.Run_Lgc("GenDta", 7, "Env=NAPROD,Brand=ESQ") Then Stop
    Case 338: If jj.Run_Lgc("GenDta", 8, "Env=NAPROD,Brand=TH") Then Stop
    Case 339: If jj.Run_Lgc("GenDta", 9, "Env=NAPROD,Brand=HB") Then Stop
    Case 340: If jj.Run_Lgc("GenDta", 10, "Env=NAPROD,Brand=Lacoste") Then Stop
    Case 341: If jj.Run_Lgc("GenDta", 11, "Env=NAPROD,Brand=Coach") Then Stop
    Case 342: If jj.Run_Lgc("GenDta", 12, "Env=NAPROD,Brand=Juicy") Then Stop
    
    Case 431: If jj.Run_Lgc("GenRpt", 1, "Env=FEPROD,Brand=ESQ") Then Stop
    Case 432: If jj.Run_Lgc("GenRpt", 2, "Env=FEPROD,Brand=TH") Then Stop
    Case 433: If jj.Run_Lgc("GenRpt", 3, "Env=FEPROD,Brand=HB") Then Stop
    Case 434: If jj.Run_Lgc("GenRpt", 4, "Env=FEPROD,Brand=Lacoste") Then Stop
    Case 435: If jj.Run_Lgc("GenRpt", 5, "Env=FEPROD,Brand=Coach") Then Stop
    Case 436: If jj.Run_Lgc("GenRpt", 6, "Env=FEPROD,Brand=Juicy") Then Stop
    Case 437: If jj.Run_Lgc("GenRpt", 7, "Env=NAPROD,Brand=ESQ") Then Stop
    Case 438: If jj.Run_Lgc("GenRpt", 8, "Env=NAPROD,Brand=TH") Then Stop
    Case 439: If jj.Run_Lgc("GenRpt", 9, "Env=NAPROD,Brand=HB") Then Stop
    Case 440: If jj.Run_Lgc("GenRpt", 10, "Env=NAPROD,Brand=Lacoste") Then Stop
    Case 441: If jj.Run_Lgc("GenRpt", 11, "Env=NAPROD,Brand=Coach") Then Stop
    Case 442: If jj.Run_Lgc("GenRpt", 12, "Env=NAPROD,Brand=Juicy") Then Stop
    End Select
Next
End Function
#End If
Function Run_Lgc_ByTrc(pTrc&) As Boolean
'Aim: Copy {pNmLgc} in [jj.Sdir_PgmLgc] to [jj.Sffn_TmpLgc] if need
'     Run {pNmLgc} in [jj.Sffn_TmpLgc] with {pTrc,pLn,pLv} as parameters.
Const cSub$ = "Run_Lgc_ByTrc"
Exit Function
E: Run_Lgc_ByTrc = True
End Function
#If Tst Then
Function Run_Lgc_ByTrc_Tst() As Boolean
Dim mCase As Byte
mCase = 2
Select Case mCase
Case 1: If jj.Run_Lgc("ExpTbl") Then Stop
Case 2: If jj.Run_Lgc("ExpFld") Then Stop
Case 9: If jj.Run_Lgc("RfhCusGp", 1, "Brand=TH") Then Stop
End Select
End Function
#End If
Function Run_Odbc(pQs$) As Boolean
Const cSub$ = "Run_Odbc"
On Error GoTo R
Dim mTrc&, mNmLgc$, mLm$: If jj.Fnd_Prm_FmTblPrm(mTrc, mNmLgc, mLm) Then ss.A 1: GoTo E
If jj.Crt_SessDta(mTrc) Then ss.A 1: GoTo E
If jj.Rfh_Lnk(mTrc) Then ss.A 2: GoTo E
Dim mFbTar$: mFbTar = jj.Sffn_SessDta(mTrc)
If jj.Run_Qry(pQs, , , , mLm, mFbTar, True) Then ss.A 3: GoTo E
If jj.Exp_SetNmtq2Mdb_ByTblOup(mFbTar) Then ss.A 4: GoTo E
Exit Function
R: ss.R
E: Run_Odbc = True: ss.B cSub, cMod, "pQs", pQs
End Function
Function Run_AcsAutoExec(oRslt As Boolean, pAcs As Access.Application) As Boolean
'Aim: Run the AutoExec in pAcs
Const cSub$ = "Run_AcsAutoExec"
On Error GoTo R
oRslt = pAcs.Eval("AutoExec()")
Exit Function
R: ss.R
E: Run_AcsAutoExec = True: ss.B cSub, cMod, "pAcs", jj.ToStr_Acs(pAcs)
X:
If jj.SysCfg_IsDbgRunAcs Then pAcs.Visible = True: Stop
End Function
Function Run_Fb(oRslt As Boolean, pFb$, pNmLgc$, pTrc& _
    , Optional pLm$ = "" _
    , Optional pHidAcs As Boolean = False) As Boolean
'Aim: Run {pNmLgc} stored as {pFb} with {pTrc,pLn,pLv} as parameters.
'Note:Each value in {pLv} is separated by |
Const cSub$ = "Run_Fb"

'Set {pNmLgc,pTrc,pLn,pLv} to {pFb}
If jj.Set_Prm(pFb, pTrc, pNmLgc, pLm) Then ss.A 1: GoTo E

'Run the pFb
Dim mAcs As Access.Application: Set mAcs = gAcs
If Not pHidAcs Then mAcs.Visible = True
On Error GoTo R
If jj.Opn_CurDb(mAcs, pFb) Then ss.A 2: GoTo E
If jj.Run_AcsAutoExec(oRslt, mAcs) Then ss.A 3, "Error in Eval('AutoExec()') in [pFb]": GoTo E
GoTo X
R: ss.R
E: Run_Fb = True: ss.B cSub, cMod, "pFbTp,pTrc,pLm,pHidAcs", pFb, pTrc, pLm
X:
'    If jj.Cls_CurDb(mAcs) Then ss.xx 3, cSub, cMod, eRunTimErr, "Error in closing CurDb", "pFb", pFb: Goto E
End Function
#If Tst Then
Function Run_Fb_Tst() As Boolean
Const cSub$ = "Run_Fb_Tst"
Const cFb1$ = "P:\WorkingDir\PgmObj\Lgc\LgcExpDb.mdb"
Const cFb2$ = "P:\WorkingDir\PgmObj\Lgc\LgcExpDb_xx.mdb"
Const cNmLgc$ = "ExpFld"
Const cTrc& = 1
Const cLp$ = ""
Const cLv$ = ""
Const cHidAcs As Boolean = True
Debug.Print Time
If jj.Run_Fb(cFb1, cNmLgc, cTrc, cLp, cLv, cHidAcs) Then Stop
'If Fb(cFb2, cNmLgc, cTrc, cLp, cLv, cHidAcs) Then Stop
Debug.Print Time
jj.Shw_Dbg cSub, cMod, "cFb1,cFb2,cNmLgc,cTrc,cLp,cLv,cHidAcs", cFb1, cFb2, cNmLgc, cTrc, cLp, cLv, cHidAcs
End Function
#End If
Function Run_Dtf(pFfnDtf$, pFfnTar$, Optional oNRec& = 0) As Boolean
'Aim:   Run {pFfnDtf}, which assume to download data to {pFfnTar} & create [mFfnFdf] & return {oNRec}
'Side Effects:
''    Delete:
''      #1 {pFfnDtf}        : will create if error
''    Create & Delete:
''      #2 [mFfnDownload]
''      #3 [mFfnDtfMsg]     : will create if error. (=*.dtf.txt
'Detail:
''       Dlt 2 files: {pFfnTar} & DirOf(pFfnTar]EndDownload
''       Build a Bat file [mFfnBat] to "#1. Run rtopcb with create *.dtf.txt" & "#2 Create EndDownload"
''       Call the bat & wait for [DirOf(pFfnTar]EndDownload] & delete [mFfnBat]
''       Import *.dtf.txt(mFfnDtfMsg) to get {oNRec}
''       create empty {pFfnTar} from [mFfnFdf], if no data is download, and if *.dtf.txt (the dtf download message) said so by return oNRec
''       If not {pKeepDtf}, Rmv {pFfnDtf} & [mFfnDtfMsg]
Const cSub$ = "Run_Dtf"
'Do build {mFfnBat}, which rtopcs {pFfnDtf}, which assume to download data to {pFfnTar}
Dim mFfnBat$, mFfnDtfMsg$, mFfnDownloadEnd$
Do
    Dim mDir$: mDir = Fct.Nam_DirNam(pFfnDtf)
    mFfnBat = mDir & "Download.Bat"
    mFfnDownloadEnd = mDir & "DownloadEnd"
    If jj.Dlt_Fil(mDir & "EndDownload") Then ss.A 2: GoTo E
    
    mFfnDtfMsg = pFfnDtf & ".txt"
    Dim mFno As Byte: If jj.Opn_Fil_ForOutput(mFno, mFfnBat, True) Then ss.A 3: GoTo E
    Print #mFno, jj.Fmt_Str("rtopcb /s ""{0}"" >""{1}""", pFfnDtf, mFfnDtfMsg)
    Print #mFno, jj.Fmt_Str("echo >""{0}""", mFfnDownloadEnd)
    Close #mFno
Loop Until True
'Do Run {mFfnBat} to download data to {pFfnTar} & with msg send to {mFfnDtfMsg}
Do
    If jj.Dlt_Fil(pFfnTar) Then ss.A 1: GoTo E
    Shell """" & mFfnBat & """", vbHide
    If Fct.WaitFor(mFfnDownloadEnd, "[" & pFfnTar & "] <--Downloading File" & vbCrLf & "[" & pFfnDtf & "] <--By Dtf File") Then ss.A 1, "User has cancelled to wait": GoTo E
    jj.Dlt_Fil mFfnBat
Loop Until True

'Find {oNRec} from {mFfnDtfMsg}
If Run_RecCnt_ByFfnDtfMsg(oNRec, mFfnDtfMsg) Then ss.A 4: GoTo E

'Do create empty {pFfnTar} from mFfnFdf, if no data is download, and if *.dtf.txt (the dtf download message) said so by return oNRec,
Dim mFfnFdf$: mFfnFdf = jj.Cut_Ext(pFfnDtf) & ".Fdf"
Do
    If VBA.Dir(pFfnTar) = "" Then
        If oNRec > 0 Then ss.A 5, "No pFfnTar is found, but NRec>0", eImpossibleReachHere: GoTo E
        Select Case Right(pFfnTar, 4)
        Case ".xls"
            Dim mWb As Workbook: If jj.Crt_Xls_FmFDF(pFfnTar, mFfnFdf) Then ss.A 6: GoTo E
        Case ".txt"
            If jj.Opn_Fil_ForOutput(mFno, pFfnTar) Then ss.A 7: GoTo E
            Close #mFno
        Case Else
            ss.A 8, "No pFfnTar is found and it is not .Xls or .Txt": GoTo E
        End Select
    End If
Loop Until True
If oNRec > 0 Then jj.Dlt_Fil pFfnDtf
Exit Function
R: ss.R
E: Run_Dtf = True: ss.B cSub, cMod, "pFfnDtf$, pFfnTar$", pFfnDtf$, pFfnTar$
End Function
#If Tst Then
Function Run_Dtf_Tst() As Boolean
Const cFfnDtf$ = "c:\tmp\aa.dtf"
Dim mNRec&: If jj.Bld_Dtf(cFfnDtf, "Select * from IIC where iclas='x12'", "192.168.103.14", , True, True, mNRec) Then Stop
End Function
#End If
Private Function Run_RecCnt_ByFfnDtfMsg(oNRec&, pFfnDtfMsg$) As Boolean
'Aim: if (CWBTF0004 No Data Mathc) or some data is download, set {oNRec}, delete pFfnDtfMsg & return OK
'     else keep {pFfnDtfMsg} & return error
Const cSub$ = "Run_RecCnt_ByFfnDtfMsg"
oNRec = 0
Dim mF As Byte: If jj.Opn_Fil_ForInput(mF, pFfnDtfMsg) Then ss.A 1: GoTo E
Do While Not EOF(mF)
    Dim mL$: Line Input #mF, mL
    If Left(mL, 9) = "CWBTF0004" Then
        Close mF
        jj.Dlt_Fil pFfnDtfMsg
        Exit Function
    End If
    If Left(mL, 17) = "Rows transferred:" Then
        oNRec = CLng(mID(mL, 18))
        If oNRec > 0 Then
            Close mF
            jj.Dlt_Fil pFfnDtfMsg
            Exit Function
        End If
    End If
Loop
Close #mF
Dim mMsg$: If jj.Read_Str_FmFil(mMsg, pFfnDtfMsg, True) Then ss.A 1: GoTo E
ss.A 2, "There is error in running the Dtf"
GoTo E
R: ss.R
E: Run_RecCnt_ByFfnDtfMsg = True: ss.B cSub, cMod, "pFfnDtfMsg,mMsg", pFfnDtfMsg, mMsg
End Function
Function Run_Prc_InMdb(pFb$, pNmPrc$, Optional pLp$, Optional p0, Optional p1, Optional p2, Optional p3, Optional p4, Optional p5) As Boolean
Const cSub$ = "Run_Prc_InMdb"
DoCmd.Hourglass True
On Error GoTo R
jj.Cls_CurDb gAcs
jj.Shw_Sts "Running Proc [" & pNmPrc & "] in [" & pFb & "]..........."
With gAcs
    .OpenCurrentDatabase pFb, True
    .Run pNmPrc, pLp, p0, p1, p2, p3, p4, p5
    .CloseCurrentDatabase
    .Quit
End With
jj.Clr_Sts
DoCmd.Hourglass False
R: ss.R
E: Run_Prc_InMdb = True: ss.B cSub, cMod, "pNmPrc,pFb,pLp,p0,p1,p2,p3,p4,p5", pNmPrc, pFb, pLp, p0, p1, p2, p3, p4, p5
X:
    DoCmd.Hourglass False
    jj.Clr_Sts
End Function
Function Run_Expr(pExpr$, Optional pAcs As Access.Application = Nothing) As Boolean
Const cSub$ = "Run_Expr"
Dim mAcs As Access.Application: Set mAcs = jj.Cv_Acs(pAcs)
On Error GoTo R
jj.Shw_Sts "Eval(" & pExpr & ") in [" & jj.ToStr_Acs(mAcs) & "] ... "
DoCmd.Hourglass True
Run_Expr = mAcs.Eval(pExpr)
GoTo X
R: ss.R
E: Run_Expr = True: ss.B cSub, cMod, "pExpr,pAcs", pExpr, jj.ToStr_Acs(pAcs)
X:
    DoCmd.Hourglass False
    jj.Clr_Sts
End Function
Function Run_Expr_InFb(pFb$, pExpr$) As Boolean
Const cSub$ = "Run_Expr_InFb"
jj.Shw_Sts "Eval(" & pExpr & ") in [" & pFb & "] ... "
DoCmd.Hourglass True
On Error GoTo R
Dim mAcs As Access.Application: Set mAcs = jj.g.gAcs
If jj.Opn_CurDb(mAcs, pFb, True) Then ss.A 1: GoTo E
Run_Expr_InFb = Run_Expr(pExpr, mAcs)
If jj.Cls_CurDb(mAcs) Then ss.A 2: GoTo E
GoTo X
R: ss.R
E: Run_Expr_InFb = True: ss.B cSub, cMod, "pExpr,pFb", pExpr, pFb
X:
    DoCmd.Hourglass False
    jj.Clr_Sts
End Function
Function Run_Expr_InFb_Tst() As Boolean
Const cSub$ = "Run_Expr_InFb_Tst"
Stop
'jj.UsrPrf_LoginAgain
Debug.Print jj.Run_Expr_InFb("tmpARCollection_1.mdb", "Rfh_InqAR()")
End Function
Function Run_Prc(pNmProc$, Optional p0, Optional p1) As Boolean
Const cSub$ = "Run_"
On Error GoTo R
Application.Run pNmProc, p0, p1
Exit Function
R: ss.R
E: Run_Prc = True: ss.B cSub, cMod, "pNmProc", pNmProc
End Function
Function Run_Qry(pNmQs$ _
    , Optional pMajBeg As Byte = 0 _
    , Optional pMajEnd As Byte = 99 _
    , Optional pSkipSetSno As Boolean = False _
    , Optional pLm$ = "" _
    , Optional pFbTar$ = "" _
    , Optional pRunOdbc As Boolean = False _
    , Optional pAcs As Access.Application = Nothing _
        ) As Boolean
Const cSub$ = "Run_Qry"
If pMajBeg > pMajEnd Then ss.A 1, "pMajEnd must >= pMajBeg and between 0 to 99": GoTo E
If pMajBeg > 99 Or pMajEnd > 99 Then ss.A 2, "pMajEnd must >= pMajBeg and between 0 to 99": GoTo E
If Left(pNmQs, 3) <> "qry" Then ss.A 3, "pNmqs must begins with qry": GoTo E
If Len(pNmQs) <= 3 Then ss.A 4, "pNmqs must >=4 chr": GoTo E

Dim mAcs As Access.Application:     Set mAcs = jj.Cv_Acs(pAcs)
Dim mDbQry As DAO.Database:         Set mDbQry = mAcs.CurrentDb
Dim mNmQsNs$: mNmQsNs = mID(pNmQs$, 4)
'If jj.Dlt_Tbl_ByPfx("tmp" & mNmQsNs, mDbQry) Then ss.A 5: GoTo E
If jj.Dlt_Tbl_ByPfx("#", mDbQry) Then ss.A 6: GoTo E
If jj.Dlt_Qry_XXNN Then ss.A 6: GoTo E
If jj.Dlt_Qry_YYNN Then ss.A 6: GoTo E

If pRunOdbc Then If jj.Bld_OdbcQs(pNmQs, pLm, , pFbTar, pRunOdbc, mDbQry) Then ss.A 7: GoTo E

Dim mAnq$(): If jj.Fnd_Anq_ByNmQs(mAnq, pNmQs, pMajBeg, pMajEnd, mDbQry) Then ss.A 8: GoTo E
If jj.Run_Qry_ByAnq(mAnq, pLm$, pFbTar, mAcs) Then ss.A 9: GoTo E
If Not pSkipSetSno Then jj.Set_TblSeqInDesc pNmQs
GoTo X
R: ss.R
E: Run_Qry = True: ss.B cSub, cMod, "pNmQs,pMajBeg,pMajEnd,pSkipSetSno,pLm,pFbTar,pRunOdbc,pAcs", pNmQs, pMajBeg, pMajEnd, pSkipSetSno, pLm, pFbTar, pRunOdbc, jj.ToStr_Acs(pAcs)
X: Set mDbQry = Nothing
End Function
#If Tst Then
Function Run_Qry_Tst() As Boolean
Const cSub$ = "Run_Qry_Tst"
If jj.Crt_SessDta(1) Then ss.A 2: GoTo E
Dim mFbTar$: mFbTar = jj.Sffn_SessDta(1)
If jj.Run_Qry("qryMPS", , , True, "Env=FEPROD,Brand=TH", mFbTar, True) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: Run_Qry_Tst = True: ss.B cSub, cMod
End Function
#End If
Function Run_Qry_ByAnq(pAnq$() _
    , Optional pLm$ = "" _
    , Optional pFbTar$ = "" _
    , Optional pAcs As Access.Application = Nothing _
    ) As Boolean
'Aim: Run {pAcs}!{pAnq} with optional parameter {pLm}
'     If there is any #Chk*, after running all pAnq$, check if #Chk* contains any record.  If yes, return error
Const cSub$ = "Run_Qry_ByAnq"
Dim N%: N = jj.Siz_Ay(pAnq): If N% = 0 Then Exit Function

Dim mAnqChk$(), mNNmqChk%: mNNmqChk = 0
Dim J%
For J = 0 To N - 1
    Dim mNmq$: mNmq = pAnq(J)
    If jj.Run_Qry_ByNmq(mNmq, pLm, pFbTar, pAcs) Then ss.A 2: GoTo E
    If InStrRev(mNmq, "_#Chk") > 0 Then
        ReDim Preserve mAnqChk$(mNNmqChk)
        mAnqChk$(mNNmqChk) = mNmq
        mNNmqChk = mNNmqChk + 1
    End If
Next
If mNNmqChk > 0 Then If jj.Shw_QryChk(mAnqChk) Then ss.A 3: GoTo E
Exit Function
R: ss.R
E: Run_Qry_ByAnq = True: ss.B cSub, cMod, "pAnq(0),pLm,pFbTar,pAcs", pAnq(0), pLm, pFbTar, jj.ToStr_Acs(pAcs)
End Function
#If Tst Then
Function Run_Qry_ByAnq_Tst() As Boolean
Const cSub$ = "Run_Qry_ByAnq_Tst"
Dim mLm$, mAnq$()
Dim mResult As Boolean
Dim mCase As Byte: mCase = 1
Dim mPfx$
Select Case mCase
Case 1
    mPfx = "qryOdbcFc_0"
    Dim mFb$: mFb = "D:\SPLHalfWayHouse\MPSDetail\VerNew@2007_01_04\WorkingDir\PgmObj\MPS_RfhFc.mdb"
    If True Then
        If jj.Crt_Tbl_FmLnkLnt(mFb, "mstEnv,mstIP,mstLib,mstBrand,tblFcPrm") Then ss.A 1: GoTo E
        If jj.Cpy_Obj_ByPfx(mPfx, acQuery, mFb) Then ss.A 1: GoTo E
    End If
    If jj.Fnd_Anq_ByPfx(mAnq, mPfx) Then ss.xx 2, cSub, cMod: Exit Function
    mLm = "Brand=ESQ,Env=FEPROD"
Case 2
    mPfx = ""
End Select
mResult = jj.Run_Qry_ByAnq(mAnq, mLm, "c:\aa.mdb")
jj.Shw_Dbg cSub, cMod, , "Result,mAnq,mLm", mResult, jj.ToStr_Ays(mAnq), mLm
Exit Function
R: ss.R
E: Run_Qry_ByAnq_Tst = True: ss.B cSub, cMod
End Function
#End If
Function Run_Qry_ByNmYY(pNmqYY$ _
    , Optional pFbTar$ = "" _
    , Optional pAcs As Access.Application = Nothing _
    ) As Boolean
'Aim: Use {pNmqYY} to build some query as *_yynn and run them.
'     {pNmqYY} is a query of name ends with _yy.  It has only one record and 2 fields of name sql, prm.  It will have as macro string to be substitued by the table as pointed by prm
Const cSub$ = "Run_Qry_ByNmYY"
On Error GoTo R
Dim mAcs As Access.Application, mDb As DAO.Database
Set mAcs = jj.Cv_Acs(pAcs): Set mDb = mAcs.CurrentDb
Dim mSqlTp$, mNmtPrm$
With mDb.QueryDefs(pNmqYY).OpenRecordset
    mSqlTp = .Fields(0).Value
    mNmtPrm = .Fields(1).Value
End With
Dim mLm$: If jj.Set_Lm_ByTbl(mLm, mNmtPrm) Then ss.A 1: GoTo E
If Run_Sql_By_Repeat_ByLm(mSqlTp, mLm, pNmqYY, pAcs) Then ss.A 1: GoTo E
GoTo X
R: ss.R
E: Run_Qry_ByNmYY = True: ss.B cSub, cMod, "pNmqYY,pFbTar,pAcs", pNmqYY, pFbTar, jj.ToStr_Acs(pAcs)
X: Set mDb = Nothing
End Function
Function Run_Qry_ByNmYY_Tst() As Boolean
If jj.Dlt_Qry_YYNN Then Stop
If jj.Run_Qry_ByNmYY("qryImpTy_02_3_AddRec_yy") Then Stop
End Function
Function Run_Qry_ByNmXX(pNmqXX$ _
    , Optional pLm$ = "" _
    , Optional pFbTar$ = "" _
    , Optional pAcs As Access.Application = Nothing _
    ) As Boolean
'Aim: Use {pNmqXX} to build some query as *_xxnn and run them.
'     {pNmqXX} is a query of name ends with _xx.  It has only one record and one field of name sql.  It will have as macro string to be substitued by {pLm}
Const cSub$ = "Run_Qry_ByNmXX"
On Error GoTo R
Dim mAcs As Access.Application, mDb As DAO.Database
Set mAcs = jj.Cv_Acs(pAcs): Set mDb = mAcs.CurrentDb
Dim mSqlTp$: mSqlTp = mDb.QueryDefs(pNmqXX).OpenRecordset.Fields(0).Value
If Run_Sql_By_Repeat_ByLm(mSqlTp, pLm, pNmqXX, pAcs) Then ss.A 1: GoTo E
GoTo X
R: ss.R
E: Run_Qry_ByNmXX = True: ss.B cSub, cMod, "pNmqxx,pLm,pFbTar,pAcs", pNmqXX, pLm, pFbTar, jj.ToStr_Acs(pAcs)
X: Set mDb = Nothing
End Function
Function Run_Qry_ByNmXX_Tst() As Boolean
Dim mCase As Byte: mCase = 2
Dim mSqlTp$, mLm$, mNmqXX$
Select Case mCase
Case 1
    mSqlTp$ = "Select ""SELECT '{Itm}' AS Itm, CByte({MaxTy}) AS MaxTy INTO [#Prm] ;"" as Sql"
    mLm = "Itm=Tbl;MaxTy=2"
    mNmqXX = "qryImpTy_01_1_Crt_xx"
    If jj.Crt_Qry(mNmqXX, mSqlTp) Then Stop
    If jj.Run_Qry_ByNmXX(mNmqXX, mLm) Then Stop
Case 2
    mSqlTp = "SELECT ""SELECT '$Ty{Itm}{N}{X}' AS Nmt Into [#Lnt] ;"" AS [Sql];"
    mLm = "Itm=Tbl;N=1,2;X=,x,xx,xx"
    mNmqXX = "qryImpTy_02_1_Crt_xx"
    If jj.Crt_Qry(mNmqXX, mSqlTp) Then Stop
    If jj.Run_Qry_ByNmXX(mNmqXX, mLm) Then Stop
End Select
End Function
Function Run_Qry_ByNm(pNmq$, Optional pAcs As Access.Application = Nothing) As Boolean
'Aim: Use {pNmq} as RunCode Function Run_name
Const cSub$ = "Run_Qry_ByNm"
'Assume the name ppp_NN_xxxx should change to ppp_xxxx
'   ppp is the prefix
'   NN  is the seq
On Error GoTo R
Debug.Print pNmq;
Dim mAcs As Access.Application: Set mAcs = jj.Cv_Acs(pAcs)
Dim mRs As DAO.Recordset
If Right(pNmq, 7) = "_RunPrc" Then
    Set mRs = mAcs.CurrentDb.QueryDefs(pNmq).OpenRecordset
    Debug.Print "<-- pAcs.Run: " & jj.ToStr_Rs(mRs)
    If jj.Run_Prc_ByRs(mRs, pAcs) Then ss.A 1: GoTo E
    GoTo X
End If

Dim mNmq$: mNmq = Replace(Replace(pNmq, "#", ""), "$", "")
Dim mNmNew$: mNmNew = jj.Fnd_FctNam_ByNmq(mNmq)
Debug.Print "<--- Renamed to "; mNmNew;
Dim mExpr$

If Right(pNmq, 4) = "_Run" Then
    Set mRs = mAcs.CurrentDb.QueryDefs(pNmq).OpenRecordset
    If mRs.EOF Then ss.A 1: GoTo E
    mExpr = mRs.Fields(0).Value
    mRs.Close
ElseIf Right(pNmq, 8) = "_RunCode" Then
    mExpr = mNmNew & "()"
Else
    ss.A 2, "pNmq must end with _RunCode or _Run or _Prc": GoTo E
End If
If jj.Run_Expr(mExpr, pAcs) Then ss.A 3: GoTo E
Debug.Print "<--- Eval OK"
Exit Function
R: ss.R
    Debug.Print "<==== Error in executing this Function"
E: Run_Qry_ByNm = True: ss.B cSub, cMod, "pNmq,pAcs,mNmNew", pNmq, jj.ToStr_Acs(pAcs), mNmNew
X:
    jj.Cls_Rs mRs
End Function
Function Run_Prc_ByRs(pRs As DAO.Recordset, Optional pAcs As Access.Application = Nothing) As Boolean
'Aim: use first field in {pRs} as NmPrc and rest as parameter to {pAcs}.Run the procedure.
Const cSub$ = "Run_Prc_ByRs"
On Error GoTo R
If pRs.EOF Then ss.A 1: GoTo E
Dim NPrm%: NPrm = pRs.Fields.Count - 1
Dim mAcs As Access.Application: Set mAcs = jj.Cv_Acs(pAcs)
With pRs
    Dim mNmPrc$: mNmPrc = .Fields(0).Value
    While Not .EOF
        If .Fields(0).Value <> mNmPrc Then ss.A 2, "The first field of pRs does not match with the first record", , "The First The Record: mNmPrc", mNmPrc: GoTo E
        If NPrm > 0 Then
            ReDim mAyV(NPrm - 1)
            Dim J%
            For J = 0 To NPrm - 1
                mAyV(J) = pRs.Fields(J + 1).Value
            Next
        End If
        Dim mRslt As Boolean
        Select Case NPrm
        Case 0: mRslt = mAcs.Run(mNmPrc)
        Case 1: mRslt = mAcs.Run(mNmPrc, mAyV(0))
        Case 2: mRslt = mAcs.Run(mNmPrc, mAyV(0), mAyV(1))
        Case 3: mRslt = mAcs.Run(mNmPrc, mAyV(0), mAyV(1), mAyV(2))
        Case 4: mRslt = mAcs.Run(mNmPrc, mAyV(0), mAyV(1), mAyV(2), mAyV(3))
        Case 5: mRslt = mAcs.Run(mNmPrc, mAyV(0), mAyV(1), mAyV(2), mAyV(3), mAyV(4))
        Case 6: mRslt = mAcs.Run(mNmPrc, mAyV(0), mAyV(1), mAyV(2), mAyV(3), mAyV(4), mAyV(5))
        Case 7: mRslt = mAcs.Run(mNmPrc, mAyV(0), mAyV(1), mAyV(2), mAyV(3), mAyV(4), mAyV(5), mAyV(6))
        Case 8: mRslt = mAcs.Run(mNmPrc, mAyV(0), mAyV(1), mAyV(2), mAyV(3), mAyV(4), mAyV(5), mAyV(6), mAyV(7))
        Case 9: mRslt = mAcs.Run(mNmPrc, mAyV(0), mAyV(1), mAyV(2), mAyV(3), mAyV(4), mAyV(5), mAyV(6), mAyV(7), mAyV(8))
        Case 10: mRslt = mAcs.Run(mNmPrc, mAyV(0), mAyV(1), mAyV(2), mAyV(3), mAyV(4), mAyV(5), mAyV(6), mAyV(7), mAyV(8), mAyV(9))
        Case 11: mRslt = mAcs.Run(mNmPrc, mAyV(0), mAyV(1), mAyV(2), mAyV(3), mAyV(4), mAyV(5), mAyV(6), mAyV(7), mAyV(8), mAyV(9), mAyV(10))
        Case 12: mRslt = mAcs.Run(mNmPrc, mAyV(0), mAyV(1), mAyV(2), mAyV(3), mAyV(4), mAyV(5), mAyV(6), mAyV(7), mAyV(8), mAyV(9), mAyV(10), mAyV(11))
        Case 13: mRslt = mAcs.Run(mNmPrc, mAyV(0), mAyV(1), mAyV(2), mAyV(3), mAyV(4), mAyV(5), mAyV(6), mAyV(7), mAyV(8), mAyV(9), mAyV(10), mAyV(11), mAyV(12))
        Case 14: mRslt = mAcs.Run(mNmPrc, mAyV(0), mAyV(1), mAyV(2), mAyV(3), mAyV(4), mAyV(5), mAyV(6), mAyV(7), mAyV(8), mAyV(9), mAyV(10), mAyV(11), mAyV(12), mAyV(13))
        Case 15: mRslt = mAcs.Run(mNmPrc, mAyV(0), mAyV(1), mAyV(2), mAyV(3), mAyV(4), mAyV(5), mAyV(6), mAyV(7), mAyV(8), mAyV(9), mAyV(10), mAyV(11), mAyV(12), mAyV(13), mAyV(14))
        Case 16: mRslt = mAcs.Run(mNmPrc, mAyV(0), mAyV(1), mAyV(2), mAyV(3), mAyV(4), mAyV(5), mAyV(6), mAyV(7), mAyV(8), mAyV(9), mAyV(10), mAyV(11), mAyV(12), mAyV(13), mAyV(14), mAyV(15))
        Case Else: ss.A 2, "Invalid # of fields in pRs": GoTo E
        End Select
        If mRslt Then ss.A 3, "Error in mAcs.Running the Rs": GoTo E
        .MoveNext
    Wend
End With
GoTo X
R: ss.R
E: Run_Prc_ByRs = True: ss.B cSub, cMod, "pRs,pAcs", jj.ToStr_Rs(pRs), jj.ToStr_Acs(pAcs)
X:
End Function
#If Tst Then
Function Run_Prc_ByRs_Tst() As Boolean
Dim mCase As Byte: mCase = 2
Dim mRs As DAO.Recordset:
Select Case mCase
Case 1
    If jj.Dlt_Tbl("#Tmp_Brk") Then Stop: GoTo E
    If jj.Crt_Tbl_FmLoFld("#Tmp", "Tbl Long,LnFld Text 255") Then Stop: GoTo E
    If jj.Run_Sql("Insert into [#Tmp] values (1,'aa,bb,cc')") Then Stop: GoTo E
    If jj.Run_Sql("Insert into [#Tmp] values (2,'xx,yy,cc')") Then Stop: GoTo E
    If jj.Crt_Qry("qryTmp", "Select 'jj.qBrkRec_ByCmd' As NmPrc,'Brk #Tmp Split LnFld Keep Tbl' As V0") Then Stop: GoTo E
    Set mRs = CurrentDb.QueryDefs("qryTmp").OpenRecordset
    If jj.Run_Prc_ByRs(mRs) Then Stop
    DoCmd.OpenTable "#Tmp_Brk"
    Stop
    If jj.Dlt_Tbl("#Tmp_Brk") Then Stop: GoTo E
    If jj.Dlt_Tbl("#Tmp") Then Stop: GoTo E
    If jj.Dlt_Qry("qryTmp") Then Stop: GoTo E
    GoTo X
Case 2
    If jj.Crt_Tbl_FmLoFld("#Tmp", "NmPrc Text 50,v0 Text 50") Then Stop: GoTo E
    If jj.Run_Sql("Insert into [#Tmp] values ('jj.Crt_Dir','C:\tmp\a1')") Then Stop: GoTo E
    If jj.Run_Sql("Insert into [#Tmp] values ('jj.Crt_Dir','C:\tdddmp\a2\')") Then Stop: GoTo E
    If jj.Run_Sql("Insert into [#Tmp] values ('jj.Crt_Dir','C:\tmp\a3\')") Then Stop: GoTo E
    If jj.Run_Sql("Insert into [#Tmp] values ('jj.Crt_Dir','C:\tmp\a4\')") Then Stop: GoTo E
    If jj.Run_Sql("Insert into [#Tmp] values ('jj.Crt_Dir','C:\tmp\a5\')") Then Stop: GoTo E
    If jj.Run_Sql("Insert into [#Tmp] values ('jj.Crt_Dir','C:\tmp\a6\')") Then Stop: GoTo E
    If jj.Run_Sql("Insert into [#Tmp] values ('jj.Crt_Dir','C:\tmp\a7\')") Then Stop: GoTo E
    If jj.Run_Sql("Insert into [#Tmp] values ('jj.Crt_Dir','C:\tmp\a8\')") Then Stop: GoTo E
    If jj.Run_Sql("Insert into [#Tmp] values ('jj.Crt_Dir','C:\tmp\a9\')") Then Stop: GoTo E
    If jj.Crt_Qry("qryTmp", "Select * from [#Tmp]") Then Stop: GoTo E
    Set mRs = CurrentDb.QueryDefs("qryTmp").OpenRecordset
    If jj.Run_Prc_ByRs(mRs) Then Stop: GoTo E
    Stop
    GoTo X
End Select
E: Run_Prc_ByRs_Tst = True
X: jj.Cls_Rs mRs
End Function
#End If
Function Run_Qry_ByNmq(pNmq$ _
    , Optional pLm$ = "" _
    , Optional pFbTar$ = "" _
    , Optional pAcs As Access.Application = Nothing _
    ) As Boolean
Const cSub$ = "Run_Qry_ByNmq"
'Aim: Run {pNmq} in {pAcs}, which may be:
'           _RunCode at endwith following characteristic:
'     #1 [{xxx@FbTar}]: {pNmq}.Sql can contain [{xxx@FbTar}]
'                              It will become [xxx] in '<<pFbTar>>' if pFbTar contains something
'                              It will become [xxx]                 if pFbTar contains nothing
'     #2 {xx}         : {pNmq}.Sql can contain {xx}
'                              It will be replaced by {pLp} & {pVayv}
If Right(pNmq, 8) = "_RunCode" Or Right(pNmq, 4) = "_Run" Or Right(pNmq, 7) = "_RunPrc" Then
    If jj.Run_Qry_ByNm(pNmq, pAcs) Then ss.A 1: GoTo E
    GoTo X
End If

If Right(pNmq, 3) = "_xx" Then
    If jj.Run_Qry_ByNmXX(pNmq, pLm, pFbTar, pAcs) Then ss.A 1: GoTo E
    GoTo X
End If
If Right(pNmq, 3) = "_yy" Then
    If jj.Run_Qry_ByNmYY(pNmq, pFbTar, pAcs) Then ss.A 1: GoTo E
    GoTo X
End If

Dim mA$: mA = Left(Right(pNmq, 5), 3)
If mA = "_xx" Or mA = "_yy" Then Debug.Print pNmq; "<--- Skipped": GoTo X

Dim mAcs As Access.Application: Set mAcs = jj.Cv_Acs(pAcs)
Dim mDbQry As DAO.Database: Set mDbQry = mAcs.CurrentDb

Dim iQry As QueryDef: Set iQry = mDbQry.QueryDefs(pNmq)

Dim iTyp As DAO.QueryDefTypeEnum: iTyp = iQry.Type
Select Case iTyp
Case DAO.QueryDefTypeEnum.dbQAction _
    , DAO.QueryDefTypeEnum.dbQAppend _
    , DAO.QueryDefTypeEnum.dbQDDL _
    , DAO.QueryDefTypeEnum.dbQDelete _
    , DAO.QueryDefTypeEnum.dbQMakeTable _
    , DAO.QueryDefTypeEnum.dbQUpdate
    Debug.Print pNmq;
    Dim mSql$: mSql = iQry.Sql
    If jj.Run_Sql(mSql, mAcs) Then ss.A 3: GoTo E
    Debug.Print
Case DAO.QueryDefTypeEnum.dbQCrosstab _
    , DAO.QueryDefTypeEnum.dbQSetOperation _
    , DAO.QueryDefTypeEnum.dbQSQLPassThrough _
    , DAO.QueryDefTypeEnum.dbQSelect
    'Do Nothing
Case Else
    ss.A 4: GoTo E
End Select
GoTo X
R: ss.R
E: Run_Qry_ByNmq = True: ss.B cSub, cMod, "pNmq,pLm,pFbTar,pAcs,TypQry,mSql", pNmq, pLm, pFbTar, jj.ToStr_Acs(pAcs), jj.ToStr_TypQry(iTyp), mSql
X: Set mDbQry = Nothing
End Function
Function Run_Qry_ByNmq_Tst() As Boolean
Const cSub$ = "Run_Qry_ByNmq_Tst"
Dim mNmq$:      mNmq = "query"
Dim mSql$:      mSql = "Select * into [{xx@FbTar}] from tmpIIC where ICLAS in ('{xx}','{yy}')"
Dim mLm$:       mLm = "xx=07,yy=57"
Dim mFbTar$:    mFbTar = "c:\aa.mdb"
Dim mCase As Byte: mCase = 3
Select Case mCase
Case 1
    ' Do create table 'tmpIIC' from 'IIC' of Dsn 'FEPROD_RBPCSF'
    If jj.Crt_Tbl_FmDSN_Nmt("FEPROD_RBPCSF", "IIC", "tmpIIC") Then Stop
Case 2
    If jj.Crt_Tbl_FmLoFld("tmpIIC", "ICLAS Text 2,Dte DATE") Then Stop
    If jj.Run_Sql("Insert into tmpIIC values ('07',Now)") Then Stop
    If jj.Run_Sql("Insert into tmpIIC values ('57',Now)") Then Stop
    If jj.Run_Sql("Insert into tmpIIC values ('17',Now)") Then Stop
    If jj.Run_Sql("Insert into tmpIIC values ('27',Now)") Then Stop
    If jj.Run_Sql("Insert into tmpIIC values ('07',Now)") Then Stop
End Select
If jj.Crt_Qry(mNmq, mSql) Then Stop
If jj.Run_Qry_ByNmq(mNmq, mLm, mFbTar) Then Stop
If jj.Run_Qry_ByNmq(mNmq, mLm) Then Stop

Dim mAcs As Access.Application: Set mAcs = jj.g.gAcs
If jj.Cls_CurDb(mAcs) Then Stop
mAcs.OpenCurrentDatabase mFbTar
mAcs.Visible = True
jj.Shw_Dbg cSub, cMod, "mNmq,mLm,mFbTar", mNmq, mLm, mFbTar
Debug.Print "-----"
Debug.Print "Examine the datatbase aa.mdb just open if the xx table contains 2 records"
End Function
Function Run_Qry_ByNmq_Tst1() As Boolean
Const cSub$ = "Run_Qry_ByNmq_Tst1"
If jj.Run_Qry_ByNmq("qryImpTy_01_1_Crt_xx", "Itm=Tbl,MaxTy=2") Then Stop
DoCmd.OpenQuery "qryImpTy_01_0_Prm"
End Function

Function Run_Qry_ByOpnQry(pNmq$) As Boolean
Const cSub$ = "Run_Qry_ByOpnQry"
On Error GoTo R
DoCmd.SetWarnings False
DoCmd.OpenQuery pNmq
GoTo X
R: ss.R
E: Run_Qry_ByOpnQry = True: ss.B cSub, cMod, "pNmq", pNmq
X: DoCmd.SetWarnings True
End Function
Function Run_Sql(pSql$, Optional pAcs As Access.Application = Nothing) As Boolean
Const cSub$ = "Run_Sql"
On Error GoTo R
Dim mAcs As Access.Application: Set mAcs = jj.Cv_Acs(pAcs)
With mAcs.DoCmd
    If Left(pSql, 6) = "Select" Then
        Dim mP%: mP = InStr(pSql, "INTO")
        Dim mA$: mA = Replace(Replace(LTrim(mID(pSql, mP + 5)), vbCr, " "), vbLf, " ")
        mP = InStr(mA, " "): If mP <> 0 Then mA = Left(mA, mP - 1)
        If jj.Dlt_Tbl(mA, mAcs.CurrentDb) Then ss.A 2, "It is create table sql, but cannot delete the table": GoTo E
    End If
    If Left(pSql, 6) = "Update" Then
        .SetWarnings False
    Else
        .SetWarnings True
    End If
    .RunSQL pSql
End With
GoTo X
Exit Function
R: ss.R
E: Run_Sql = True: ss.B cSub, cMod, "pSql,pAcs", pSql, jj.ToStr_Acs(pAcs)
   Debug.Print pSql
X:
Exit Function
If jj.Crt_Qry("#Debug", pSql, mAcs.CurrentDb) Then Stop
mAcs.DoCmd.OpenQuery "#Debug", acViewDesign
If Not mAcs.Visible Then mAcs.Visible = True
Stop
End Function
#If Tst Then
Function Run_Sql_Tst() As Boolean
If jj.Crt_Tbl_FmLoFld("#Tmp", "aa Text 1") Then Stop
If jj.Run_Sql("Select 234,324 into [#Tmp] ") Then Stop
End Function
#End If
Function Run_Sql_ByDbExec(pSql$, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Run_Sql_ByDbExec"
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
On Error GoTo R
mDb.Execute pSql
'oNRecAffected = mDb.RecordsAffected
Exit Function
R: ss.R
E: Run_Sql_ByDbExec = True: ss.B cSub, cMod, "pSql,pDb", pSql, jj.ToStr_Db(pDb)
End Function

