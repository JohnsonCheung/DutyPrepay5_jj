Attribute VB_Name = "xCv"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xCv"
Function Cv_AyV(pAyV) As String()
Dim mT$: mT = TypeName(pAyV)
If mT = "Empty" Then: Exit Function
If mT = "String()" Then Cv_AyV = pAyV: Exit Function
If mT = "Variant()" Then
    Dim J%, mN%: mN = jj.Siz_Ay(pAyV)
    ReDim mA$(mN - 1)
    Dim mAyV(): mAyV = pAyV
    For J = 0 To mN - 1
        mA(J) = mAyV(J)
    Next
    Cv_AyV = mA
    Exit Function
End If
ReDim mA$(0): mA(0) = pAyV
Cv_AyV = mA
End Function
#If Tst Then
Function Cv_AyV_Tst() As Boolean
Dim mB$():
Dim mA$()
Dim mX
mB = Cv_AyV(mX)
mB = Cv_AyV(mA)

mA = Split("1 2 3")
mB = Cv_AyV(mA)

Dim mV(): mV = Array(1, 2, "3", Date)
mB = Cv_AyV(mV)
End Function
#End If
Function Cv_Lp16vToAm(oAm() As tMap, pLp$, pV0 _
    , Optional pV1 _
    , Optional pV2 _
    , Optional pV3 _
    , Optional pV4 _
    , Optional pV5 _
    , Optional pV6 _
    , Optional pV7 _
    , Optional pV8 _
    , Optional pV9 _
    , Optional pV10 _
    , Optional pV11 _
    , Optional pV12 _
    , Optional pV13 _
    , Optional pV14 _
    , Optional pV15 _
    ) As Boolean
Const cSub$ = "Cv_Lp16vToAm"
If pLp = "" Then Clr_Am oAm: Exit Function
Dim mAn$(), mN%: If jj.Brk_Ln2Ay(mAn, pLp) Then ss.A 1: GoTo E
mN = jj.Siz_Ay(mAn): If mN > 15 Then ss.A 2, "too much (>15) itm in pLp": GoTo E
ReDim oAm(mN - 1)
Dim J%
For J = 0 To mN - 1
    oAm(J).F1 = mAn(J)
    Select Case J
    Case 0: oAm(J).F2 = pV0
    Case 1: oAm(J).F2 = pV1
    Case 2: oAm(J).F2 = pV2
    Case 3: oAm(J).F2 = pV3
    Case 4: oAm(J).F2 = pV4
    Case 5: oAm(J).F2 = pV5
    Case 6: oAm(J).F2 = pV6
    Case 7: oAm(J).F2 = pV7
    Case 8: oAm(J).F2 = pV8
    Case 9: oAm(J).F2 = pV9
    Case 10: oAm(J).F2 = pV10
    Case 11: oAm(J).F2 = pV11
    Case 12: oAm(J).F2 = pV12
    Case 13: oAm(J).F2 = pV13
    Case 14: oAm(J).F2 = pV14
    Case 15: oAm(J).F2 = pV15
    End Select
Next
Exit Function
E: Cv_Lp16vToAm = True: ss.B cSub, cMod, "pLp,pV0,..", pLp, pV0, ".."
End Function
Function Cv_Ays2AyV(pAys$()) As Variant()
Dim N%: N% = jj.Siz_Ay(pAys): If N% = 0 Then Exit Function
ReDim mAyV(N - 1)
Dim J%
For J = 0 To N - 1
    mAyV(J) = pAys(J)
Next
Cv_Ays2AyV = mAyV
End Function
Function Cv_Col2Nxt$(pCol$, pNCol As Byte)
Const cSub$ = "Cv_Col2Nxt"
If pNCol = 0 Then Cv_Col2Nxt$ = pCol: Exit Function
If Len(pCol) = 1 Then
    If pCol = "Z" Then
        Cv_Col2Nxt = "AA"
        Exit Function
    End If
    Cv_Col2Nxt = Chr(Asc(pCol) + 1)
    Exit Function
End If
If Len(pCol) <> 2 Then ss.A 1, "pCol must be 1 or 2 char": GoTo E
If Right(pCol, 1) = "Z" Then Cv_Col2Nxt = Chr(Asc(Left(pCol, 1)) + 1) & "A": Exit Function
Cv_Col2Nxt = Left(pCol, 1) & Chr(Asc(Right(pCol, 1)) + 1)
Exit Function
E: ss.B cSub, cMod, "pCol,pNCol", pCol, pNCol
End Function
Function Cv_Bool$(pIfTrue As Boolean, pTruePart$)
If pIfTrue Then Cv_Bool = pTruePart: Exit Function
End Function
Function Cv_Str_ByQ$(pV, pQ$)
If Nz(pV, "") = "" Then Exit Function
Cv_Str_ByQ = jj.Q_S(pV, pQ)
End Function
Function Cv_Str$(pV, Optional pPfx$ = "", Optional pSfx$ = "", Optional pQ As Boolean = False)
If Nz(pV, "") = "" Then Exit Function
If pQ Then Cv_Str = pPfx & jj.Q_V(pV) & pSfx: Exit Function
Cv_Str = pPfx & pV & pSfx
End Function
#If Tst Then
Function Cv_Str_Tst() As Boolean
Debug.Print jj.Cv_Str("sdfdf", "=", , True)
End Function
#End If
Function Cv_GpBy$(pGpBy$)
If pGpBy = "" Then Exit Function
Cv_GpBy = " Group by " & pGpBy
End Function
Function Cv_OrdBy$(pOrdBy$)
If pOrdBy = "" Then Exit Function
Cv_OrdBy = " Order by " & pOrdBy
End Function
Function Cv_Where$(pWhere$)
If pWhere = "" Then Exit Function
Cv_Where = " Where " & pWhere
End Function
Function Cv_Ws(pWs As Worksheet) As Worksheet
If TypeName(pWs) = "Nothing" Then Set Cv_Ws = jj.g.gXls.ActiveSheet: Exit Function
Set Cv_Ws = pWs
End Function
Function Cv_Prj(oPrj As VBProject, pNmPrj$, Optional pAcs As Access.Application = Nothing) As Boolean
Const cSub$ = "Prj"
Dim mAcs As Access.Application: Set mAcs = jj.Cv_Acs(pAcs)
If pNmPrj = "" Then Set oPrj = mAcs.VBE.ActiveVBProject: Exit Function
If jj.Fnd_Prj(oPrj, pNmPrj, mAcs) Then ss.A 2: GoTo E
Exit Function
E: Cv_Prj = True: ss.B cSub, cMod, "pNmPrj", pNmPrj
End Function
Function Cv_ImpWsToTbl(pFxImp$, Optional pSetWs$ = "*", Optional pFb$ = "") As Boolean
'Aim: Cv pFxImp!pSetWs in "Import" format to table(s) in {pFb}.  The table name will [>{NmWs}]
Const cSub$ = "ImpWsToTbl"
On Error GoTo R
Dim mDb As DAO.Database, mAcs As Access.Application

If pFb <> "" Then If VBA.Dir(pFb) = "" Then If jj.Crt_Fb(pFb, True) Then ss.A 1: GoTo E
If jj.Cv_Db_FmFb(mDb, pFb) Then ss.A 2: GoTo E
If jj.Crt_Tbl_ForTxtSpec(mDb) Then ss.A 4: GoTo E
If pFb <> "" Then mDb.Close

If jj.Cv_Acs_FmFb(mAcs, pFb) Then ss.A 3: GoTo E
Set mDb = mAcs.CurrentDb

Dim mWb As Workbook: If jj.Opn_Wb_R(mWb, pFxImp) Then ss.A 1: GoTo E
Dim mAnWs$(): If jj.Fnd_AnWs_BySetWs(mAnWs, mWb, pSetWs) Then ss.A 2: GoTo E
Dim J%, mNWs%: mNWs = jj.Siz_Ay(mAnWs)
ReDim mAyFfn$(mNWs - 1)
Dim mXls As Excel.Application: Set mXls = mWb.Application: mXls.DisplayAlerts = False

For J = 0 To mNWs - 1
    Dim mWs As Worksheet: Set mWs = mWb.Sheets(mAnWs(J))
    Dim mAmFld() As tMap, mAyCno() As Byte: If jj.Fnd_AyCnoImpFld(mAyCno, mAmFld, mWs.Range("A5")) Then ss.A 1: GoTo E
    If jj.Clr_ImpWs(mWs.Range("A5")) Then ss.A 1: GoTo E
    
    'Save to Csv
    mAyFfn(J) = "c:\tmp\Exp_Ws2Tbl" & jj.Fct.TimStmp & "_" & mAnWs(J) & ".csv"
    mWs.SaveAs mAyFfn(J), Excel.XlFileFormat.xlCSVWindows
    
    'Crt_TxtSpec_Fix
    If jj.Crt_TxtSpec_Delimi(">" & mAnWs(J), mAmFld, mDb) Then ss.A 1: GoTo E
    If jj.Crt_Tbl_FmAmFld(">" & mAnWs(J), mAmFld, mDb) Then ss.A 1: GoTo E
Next
jj.Cls_Wb mWb, False, True
For J = 0 To mNWs - 1
    mAcs.DoCmd.TransferText acImportDelim, ">" & mAnWs(J), ">" & mAnWs(J), mAyFfn(J), True
    If jj.Dlt_Fil(mAyFfn(J)) Then ss.A 2: GoTo E
Next
GoTo X
R: ss.R
E: Cv_ImpWsToTbl = True: ss.B cSub, cMod, "pFxImp,pSetWs", pFxImp, pSetWs
X: If pFb <> "" Then Set mDb = Nothing: jj.Cls_CurDb mAcs
   On Error Resume Next
   jj.Cls_Wb mWb, False, True
   mXls.DisplayAlerts = True
   jj.Clr_Sts
End Function
#If Tst Then
Function Cv_ImpWsToTbl_Tst() As Boolean
Const cNmWs$ = "TblL"
If True Then If jj.Cpy_Fil("p:\AppDef_Meta\MetaDb.xls", "c:\tmp\aa.xls", True) Then Stop: GoTo E
If Cv_ImpWsToTbl("c:\tmp\aa.xls", , "C:\Tmp\aa.mdb") Then Stop: GoTo E
Stop
Exit Function
E: Cv_ImpWsToTbl_Tst = True
X:
End Function
#End If
Function Cv_ImpWsToXls(pFxTar$, pFxImp$, pSetWs$) As Boolean
'Aim: Cv pFxImp!pSetWs in "Import" format to pFxTar.
'     If ws exists in {pFxTar}, the ws will be replaced and position will be retain, else ws will be added to end.
Const cSub$ = "ImpWsToXls"
On Error GoTo R
Dim mWbFm As Workbook
If jj.Opn_Wb_R(mWbFm, pFxImp) Then ss.A 2: GoTo E
Dim mAnWs$(): If jj.Fnd_AnWs_BySetWs(mAnWs, mWbFm, pSetWs) Then ss.A 3: GoTo E
Dim iWs As Worksheet, J%
For J = 0 To jj.Siz_Ay(mAnWs) - 1
    Set iWs = mWbFm.Sheets(mAnWs(J))
    If jj.Clr_ImpWs(iWs) Then ss.A 3: GoTo E
Next
If pSetWs = "*" Then
    If jj.Dlt_Fil(pFxTar) Then ss.A 4: GoTo E
    mWbFm.SaveAs pFxTar
    GoTo X
End If

Dim mWbTo As Workbook, mToBeDelete As Boolean
If VBA.Dir(pFxTar) = "" Then
    If jj.Crt_Wb(mWbTo, pFxTar) Then ss.A 1: GoTo E
    mToBeDelete = True
Else
    If jj.Opn_Wb_RW(mWbTo, pFxTar) Then ss.A 2: GoTo E
    mToBeDelete = False
End If
    
For J = 0 To jj.Siz_Ay(mAnWs) - 1
    If jj.Repl_Ws_In2Wb(mWbTo, mWbFm, mAnWs(J)) Then ss.A 4: GoTo E
Next
If mToBeDelete Then If jj.Dlt_Ws_InWb(mWbTo, "ToBeDelete") Then ss.A 5: GoTo E
GoTo X
R: ss.R
E: Cv_ImpWsToXls = True: ss.B cSub, cMod, "pFxImp,pSetWs,pFxTar", pFxImp, pSetWs, pFxTar
X: jj.Cls_Wb mWbTo, True, True
   jj.Cls_Wb mWbFm, False, True
End Function
#If Tst Then
Function Cv_ImpWsToXls_Tst() As Boolean
Dim mFxImp$, mSetWs$, mFxTar$
Dim mCase As Byte: mCase = 2
Select Case mCase
Case 1
    mFxImp = "P:\AppDef_Meta\MetaDb.xls"
    mSetWs = "Tbl,Ele,Fmt,Gp,GpF,Mdb,Schm,MdbS"
    mFxTar = "c:\tmp\tmpJMtcDb.xls"
Case 2
    mFxImp = "P:\AppDef_Meta\MetaPgm.xls"
    mSetWs = "Ass"
    mFxTar = "c:\tmp\tmpJMtcPgm.xls"
End Select
If Cv_ImpWsToXls(mFxTar, mFxImp, mSetWs) Then Stop: GoTo E
If jj.Crt_Tbl_FmLnkSetWs(mFxTar, mSetWs, ">") Then Stop: GoTo E
Exit Function
E: Cv_ImpWsToXls_Tst = True
End Function
#End If
Function Cv_TypDAO2QChr$(pTypDAO As DAO.DatabaseTypeEnum)
Cv_TypDAO2QChr = Cv_TypSim2QChr(jj.Cv_TypDAO2Sim(pTypDAO))
End Function
Function Cv_TypSim2QChr$(pTypSim As eTypSim)
Select Case pTypSim
Case eTypSim_Str: Cv_TypSim2QChr = cQSng
Case eTypSim_Dte: Cv_TypSim2QChr = "#"
Case Else: Cv_TypSim2QChr = ""
End Select
End Function
Function Cv_Fb2InFb$(pFb$)
If pFb = "" Then Exit Function
Cv_Fb2InFb = " in '" & pFb & cQSng
End Function
Function Cv_LnFld_FmFdf(oLnFld$, pFfnFdf$) As Boolean
Const cSub$ = "LnFld_FmFdf"
'Aim: Cv {pFfnFdf} to oLnFld.  If will be either * or [list of field] as yymd_ prefix field converted
'FDF format:
'PCFDF
'PCFT 1
'PCFO 1,1,5,1,1
'PCFL IID 1 2
'PCFL IPROD 1 15
Dim mFno As Byte: If jj.Opn_Fil_ForInput(mFno, pFfnFdf) Then ss.A 1: GoTo E
Dim mL$, mA$
Line Input #mFno, mL: mA = "PCFDF": If mL <> mA Then ss.A 2, "Line 1 of Fdf must be [" & mA & "]", "Line", mL: GoTo E
Line Input #mFno, mL: mA = "PCFT 16": If mL <> mA And mL <> "PCFT 1" Then ss.A 2, "Line 2 of Fdf must be [" & mA & "] or [PCFT 1", "Line", mL: GoTo E
Line Input #mFno, mL: mA = "PCFO 1,1,5,1,1": If mL <> mA Then ss.A 3, "Line 3 of Fdf must be [" & mA & "]", "Line", mL: GoTo E
oLnFld = ""
While Not EOF(mFno)
    Line Input #mFno, mL: mA = "PCFL ": If Left(mL, 5) <> mA Then ss.A 4, "Line 4 onward of Fdf must begin with [" & mA & "]", "Line", mL: GoTo E
    Dim mX$(): mX = Split(mL)
    oLnFld = jj.Add_Str(oLnFld, mX(1))
Wend
Close #mFno
oLnFld = Cv_LnFld(oLnFld)
Exit Function
R: ss.R
E: Cv_LnFld_FmFdf = True: ss.B cSub, cMod, "pFfnFdf", pFfnFdf
End Function
#If Tst Then
Function Cv_LnFld_FmFdf_Tst() As Boolean
Const cFfnDtf$ = "c:\aa.dtf"
Const cFfnFdf$ = "c:\aa.fdf"
If jj.Bld_Dtf(cFfnDtf, "Select IIC.*,ICUMPC AS yymd_ICUMPC  from IIC", "192.168.103.14", , , True) Then Stop
Dim mLnFld$: If Cv_LnFld_FmFdf(mLnFld, cFfnFdf) Then Stop
Debug.Print mLnFld
End Function
#End If
Function Cv_LnFld$(pLnFld$)
If InStr(pLnFld, "yymd_") <= 0 Then Cv_LnFld = "*": Exit Function
'
Dim mLnFld$
Dim mAy$(): mAy = Split(pLnFld, cComma)
Dim I%: For I = 0 To jj.Siz_Ay(mAy) - 1
    If Left(mAy(I), 5) = "yymd_" Then
        mLnFld = jj.Add_Str(mLnFld, jj.Fmt_Str("Cdate(IIf(yymd_{0}=0,0,IIf(yymd_{0}=99999999,'9999/12/31',format(yymd_{0},'0000\/00\/00')))) as {0}", mID(mAy(I), 6)))
    Else
        mLnFld = jj.Add_Str(mLnFld, mAy(I))
    End If
Next
Cv_LnFld = mLnFld
End Function
#If Tst Then
Function Cv_LnFld_Tst() As Boolean
Debug.Print jj.Cv_LnFld("yymd_abc,xyz")
jj.Shw_DbgWin
End Function
#End If
Function Cv_Db(pDb As DAO.Database) As DAO.Database
Static xDb As DAO.Database
If TypeName(pDb) <> "Nothing" Then Set Cv_Db = pDb: Exit Function
If jj.IsAcs Then Set Cv_Db = CurrentDb: Exit Function
If TypeName(xDb) <> "Nothing" Then Set Cv_Db = xDb: Exit Function
Const cFb$ = "c:\tmp\DbUsedByXls.mdb"
If VBA.Dir(cFb) = "" Then
    Set xDb = jj.g.gDbEng.CreateDatabase(cFb, DAO.dbLangGeneral)
Else
    Set xDb = jj.g.gDbEng.OpenDatabase(cFb)
End If
Set Cv_Db = xDb
End Function
#If Tst Then
Function Cv_Db_Tst() As Boolean
Debug.Print Application.DBEngine.Workspaces(0).Databases.Count
Dim mDb As DAO.Database
Set mDb = jj.Cv_Db(mDb)
Debug.Print Application.DBEngine.Workspaces(0).Databases.Count
Stop
End Function
#End If
Function Cv_Acs_FmFb(oAcs As Access.Application, pFb$) As Boolean
Const cSub$ = "Acs_FmFb"
If pFb = "" Then Set oAcs = Application: Exit Function
If VBA.Dir(pFb) = "" Then If jj.Crt_Fb(pFb) Then ss.A 1: GoTo E
Set oAcs = jj.g.gAcs
If jj.Opn_CurDb(oAcs, pFb) Then ss.A 2: GoTo E
Exit Function
E: Cv_Acs_FmFb = True: ss.B cSub, cMod, "pFb", pFb
End Function
Function Cv_App(pApp As Application) As Application
Static xApp As Application
If TypeName(pApp) <> "Nothing" Then Set Cv_App = pApp: Exit Function
Set Cv_App = Application
End Function
Function Cv_Acs(pAcs As Access.Application) As Access.Application
Static xAcs As Access.Application
If TypeName(pAcs) <> "Nothing" Then Set Cv_Acs = pAcs: Exit Function
If jj.IsAcs Then Set Cv_Acs = Application: Exit Function
If TypeName(xAcs) <> "Nothing" Then Set Cv_Acs = xAcs: Exit Function
Set xAcs = New Access.Application
Set Cv_Acs = xAcs
Const cFb$ = "c:\tmp\UsedByCvAcsInXls.mdb"
If VBA.Dir(cFb) = "" Then jj.Crt_Fb cFb
xAcs.OpenCurrentDatabase cFb
End Function
Function Cv_Db_FmFb(oDb As DAO.Database, pFb$) As Boolean
Const cSub$ = "Db_FmFb"
If pFb = "" Then Set oDb = CurrentDb: Exit Function
If VBA.Dir(pFb) = "" Then
    If jj.Crt_Db(oDb, pFb) Then ss.A 1: GoTo E
    Exit Function
End If
If jj.Opn_Db_RW(oDb, pFb) Then ss.A 2: GoTo E
Exit Function
R: ss.R
E: Cv_Db_FmFb = True: ss.B cSub, cMod, "pFb", pFb
End Function
Function Cv_Fdf2Schema(pFfnFdf$) As Boolean
Const cSub$ = "Fdf2Schema"
'Aim: Create [Schema.ini] in the same directory as {pFfnFdf}
'FDF format:
'PCFDF
'PCFT 1
'PCFO 1,1,1,1,1
'PCFL IID 1 2
'PCFL IPROD 1 15
'Schema.ini Format:
'[IIM_Short.txt]
'ColNameHeader = False
'Format = FixedLength
'MaxScanRows = 100
'CharacterSet = OEM
'Col1="IID" Char Width 2
'Col2="IPROD" Char Width 15
Dim mFnn$, mDir$, mExt$: If jj.Brk_Ffn_To3Seg(mDir, mFnn, mExt, pFfnFdf) Then ss.A 1: GoTo E
Dim mFfnSch$: mFfnSch$ = mDir & "Schema.ini"
Dim mF_I As Byte: If jj.Opn_Fil_ForInput(mF_I, pFfnFdf) Then ss.A 2: GoTo E
Dim mF_O As Byte: If jj.Opn_Fil_ForOutput(mF_O, mFfnSch, True) Then ss.A 3: GoTo E
Dim mL$, mA$
mA = "PCFDF":           Line Input #mF_I, mL: If mL <> mA Then ss.A 4, "Line must be [" & mA & "]", ePrmErr, "Line#1", mL: GoTo E
mA = "PCFT 1":          Line Input #mF_I, mL: If mL <> mA Then ss.A 5, "Line must be [" & mA & "]", ePrmErr, "Line#1", mL: GoTo E
mA = "PCFO 1,1,5,1,1":  Line Input #mF_I, mL: If mL <> mA Then ss.A 6, "Line must be [" & mA & "]", ePrmErr, "Line#1", mL: GoTo E
Print #mF_O, "[" & mFnn & ".txt]"
Print #mF_O, "ColNameHeader = False"
Print #mF_O, "Format = FixedLength"
Print #mF_O, "MaxScanRows = 100"
Print #mF_O, "CharacterSet = OEM"
Dim N As Byte
While Not EOF(mF_I)
    Line Input #mF_I, mL: N = N + 1
    Dim mNm$, mT$, mW% 'Name, Type, Width
    Dim mB$(): mB = Split(mL)
    If mB(0) <> "PCFL" Then ss.A 7, "Line#" & N & " does not begin with PCFL", ePrmErr, "Line", mL: GoTo E
    mNm = mB(1): mW = Val(mB(3))
    Select Case mB(2)
    Case "1": mT = "Char"
    Case "2"
        If InStr(mB(3), "/") > 0 Then
            mT = "Double"
        Else
            Select Case mW
            Case Is <= 2: mT = "Byte"
            Case Is <= 4: mT = "Integer"
            Case Is <= 9: mT = "Long"
            Case Else
                mT = "Double"
            End Select
        End If
    End Select
    Print #mF_O, jj.Fmt_Str("Col{0}=""{1}"" {2} Width {3}", N, mNm, mT, mW)
'   Print #mF_O, "Col" & N & "=""" & mNm & """ " & mT & " Width " & mW
Wend
Exit Function
R: ss.R
E: Cv_Fdf2Schema = True: ss.B cSub, cMod, "pFfnFdf", pFfnFdf
X: Close #mF_I, #mF_O
End Function
Function Cv_Fdf2Schema_Tst() As Boolean
If jj.Cv_Fdf2Schema("D:\Data\Johnson Cheung\MyDoc\My Projects\My Projects Library\Ldb\Ldb\WorkingDir\Data\IIM.fdf") Then Stop
End Function
Function Cv_AyByt2LoCol$(pAyByt() As Byte)
Dim N%: N = jj.Siz_Ay(pAyByt)
Dim mA$
Dim J%: For J = 0 To N - 1
    mA = jj.Add_Str(mA, jj.Cv_Cno2Col(pAyByt(J)), " ")
Next
Cv_AyByt2LoCol = mA
End Function
Function Cv_Cno2Col$(pCno As Byte)
Dim A As Byte:      A = Asc("A")
Dim mA1 As Byte:    mA1 = (pCno - 1) \ 26
Dim mA2 As Byte:    mA2 = (pCno - 1) Mod 26
If mA1 = 0 Then Cv_Cno2Col = Chr(A + mA2): Exit Function
Cv_Cno2Col = Chr(A - 1 + mA1) & Chr(A + mA2)
End Function
Function Cv_Col2Cno(pCol$) As Byte
Cv_Col2Cno = 0
Dim mC1$
Dim mC2$
If Len(pCol) = 1 Then
    mC1 = UCase(pCol)
    If "A" > mC1 Or mC1 > "Z" Then Exit Function
    Cv_Col2Cno = Asc(mC1) - 64
    Exit Function
End If
If Len(pCol) = 2 Then
    mC1 = UCase(Left(pCol, 1))
    mC2 = UCase(Right(pCol, 1))
    If "A" > mC1 Or mC1 > "Z" Then Exit Function
    If "A" > mC2 Or mC2 > "Z" Then Exit Function
    Cv_Col2Cno = 26 * (Asc(mC1) - 64) + Asc(mC2) - 64
End If
End Function
Function Cv_Coll2Ay(pColl As VBA.Collection, oNItm%, oAys$(), Optional pTrim As Boolean = False) As Boolean
Const cSub$ = "Coll2Ay"
oNItm = 0
If jj.IsNothing(pColl) Then Cv_Coll2Ay = True: Exit Function
If pColl.Count = 0 Then Cv_Coll2Ay = True: Exit Function

oNItm = pColl.Count
ReDim oAys(0 To oNItm - 1)
Dim J%
On Error GoTo R
If pTrim Then
    For J = 0 To oNItm - 1
        oAys(J) = Trim(pColl(J + 1))
    Next
Else
    For J = 0 To oNItm - 1
        oAys(J) = pColl(J + 1)
    Next
End If
Exit Function
R: ss.R
E: Cv_Coll2Ay = True: ss.B cSub, cMod
End Function
Function Cv_CollKvStr_To2Ay(pColl As VBA.Collection, oKey$(), oVal$(), Optional pKeyQ$ = "") As Boolean
'Aim: Convert a Collection of string of format {Key}={Val} in Array {oKey} & {oVal}
Dim mNColl%: mNColl = jj.Siz_Coll(pColl)
If mNColl = 0 Then Exit Function
ReDim oKey$(0 To mNColl - 1), oVal$(0 To mNColl - 1)
Dim J As Byte, mPos As Byte
Dim MM: For Each MM In pColl
    mPos = InStr(MM, "=")
    If mPos = 0 Then
        oKey(J) = jj.Q_S(CStr(MM), pKeyQ)
    Else
        oKey(J) = jj.Q_S(Left(MM, mPos - 1), pKeyQ)
        oVal(J) = mID$(MM, mPos + 1)
    End If
    J = J + 1
Next
End Function
Function Cv_CollKvStr_To2Ay_Tst() As Boolean
Dim mKeyLst$, mValLst$, mColl As VBA.Collection
Dim AyK$(), AyV$()
Cv_CollKvStr_To2Ay mColl, AyK, AyV
mKeyLst = Join(AyK, cComma)
mValLst = Join(AyV, cComma)
jj.Cmp_TwoTbl "Table1", "Table2", "Id", "aa,bb"
End Function
Function Cv_Csv2Xls(pFfnCsv$, Optional pFx$ = "", Optional pOvrWrt As Boolean = False, Optional pKeepCsv = False) As Boolean
Const cSub$ = "Csv2Xls"
'Aim: Cv {pFfnCsv} to {pFx}
If VBA.Dir(pFfnCsv) = "" Then ss.A 1, "{pFfnCsv} not exist": GoTo E
If pFx = "" Then pFx = jj.Repl_Ext(pFfnCsv, ".xls")
If jj.Ovr_Wrt(pFx, pOvrWrt) Then ss.A 1: GoTo E
Dim mWb As Workbook: If jj.Opn_Wb_R(mWb, pFfnCsv) Then ss.A 2: GoTo E
mWb.SaveAs pFx, XlFileFormat.xlWorkbookNormal
mWb.Close
If Not pKeepCsv Then jj.Dlt_Fil pFfnCsv
Exit Function
R: ss.R
E: Cv_Csv2Xls = True: ss.B cSub, cMod, "pFfnCsv,pFx,pOvrWrt,pKeepCsv", pFfnCsv, pFx, pOvrWrt, pKeepCsv
End Function
#If Tst Then
Function Cv_Csv2Xls_Tst() As Boolean
Const cFfnCsv$ = "c:\aa.csv"
Const cFx$ = "c:\aa.xls"
Close
Dim mFno As Byte: If jj.Opn_Fil_ForOutput(mFno, cFfnCsv, True) Then Stop
Print #mFno, "aa,bb,cc,dd"
Print #mFno, "1,23,4,6"
Print #mFno, "2,3223,1234,1/2/2007"
Print #mFno, "221,323,423,621"
Close #mFno
If jj.Cv_Csv2Xls(cFfnCsv, , True) Then Stop
Dim mWb As Workbook: If jj.Opn_Wb(mWb, cFx) Then Stop
mWb.Application.Visible = True
End Function
#End If
Function Cv_Dte2Fy$(Optional pDte As Date = 0)
Cv_Dte2Fy = jj.ToStr_FYNo(Cv_Dte2FyNo(pDte))
End Function
Function Cv_Dte2FyNo(Optional pDte As Date = 0) As Byte
Dim mDte As Date: mDte = IIf(pDte = 0, Date, pDte)
If Month(mDte) = 1 Then Cv_Dte2FyNo = Year(mDte) - 2000: Exit Function
Cv_Dte2FyNo = Year(mDte) - 1999
End Function
Function Cv_Dte2Qtr(Optional pDte As Date = 0) As String
Dim mDte As Date: mDte = IIf(pDte = 0, Date, pDte)
Dim mMM%: mMM = Month(mDte)
If mMM = 1 Then
    Cv_Dte2Qtr = "Q4"
Else
    Cv_Dte2Qtr = "Q" & Int((mMM - 2) / 3) + 1
End If
End Function
Function Cv_Rge2Sq(oSq As tSq, pRge As Range) As Boolean
Const cSub$ = "Rge2Sq"
On Error GoTo R
With pRge
    oSq.c1 = .Column
    oSq.c2 = .Column + .Columns.Count - 1
    oSq.r1 = .Row
    oSq.r2 = .Row + .Rows.Count - 1
End With
Exit Function
R: ss.R
E: Cv_Rge2Sq = True: ss.B cSub, cMod, "Rge", jj.ToStr_Rge(pRge)
    
End Function
Function Cv_RnoColRge2Adr$(pColRge$, pRow&, Optional pNRow& = 1)
Dim p As Byte: p = InStr(pColRge, ":")
If p = 0 Then Cv_RnoColRge2Adr = pColRge & pRow & ":" & pColRge & pRow + pNRow - 1: Exit Function
Cv_RnoColRge2Adr = Left(pColRge, p - 1) & pRow & ":" & mID(pColRge, p + 1) & pRow + pNRow - 1
End Function
Function Cv_Rs2Am(oAm() As tMap, pRs As DAO.Recordset) As Boolean
Dim J As Byte
ReDim oAm(0 To pRs.Fields.Count - 1)
Dim iFld As DAO.Field: For Each iFld In pRs.Fields
    With oAm(J)
        .F1 = iFld.Name
        .F2 = Nz(iFld.Value, "")
    End With
    J = J + 1
Next
End Function
Function Cv_Sq2NCol(pSq As tSq) As Byte
Const cSub$ = "Sq2NCol"
On Error GoTo R
With pSq
    Cv_Sq2NCol = .c2 - .c1 + 1
End With
R: ss.R
E: ss.B cSub, cMod, "pSq", jj.ToStr_Sq(pSq)
End Function
Function Cv_Sq2Rge(oRge As Range, pWs As Worksheet, pSq As tSq) As Boolean
Const cSub$ = "Sq2Rge"
On Error GoTo R
With pSq
    Set oRge = pWs.Range(pWs.Cells(.r1, .c1), pWs.Cells(.r2, .c2))
End With
Exit Function
R: ss.R
E: Cv_Sq2Rge = True: ss.B cSub, cMod, "pWs,pSq", jj.ToStr_Ws(pWs), jj.ToStr_Sq(pSq)
End Function
Function Cv_Tbl2Tbl(pNmtFm$, pNmtTo$, pLm$) As Boolean
'Aim: Transform {pNmtFm} into {pNmtTo} by using {pLm}
Const cSub$ = "Tbl2Tbl"
On Error GoTo R
If Not jj.IsTbl(pNmtFm) Then ss.A 1, "pNmtFm not exist", , "pNmtFm,pNmtTo", pNmtFm, pNmtTo
If jj.Dlt_Tbl(pNmtTo) Then ss.A 2: GoTo E
Dim mAm() As tMap: mAm = jj.Get_Am_ByLm(pLm)
Dim mAyFm$(): If jj.Cpy_AmF1_ToAy(mAyFm, mAm) Then ss.A 4: GoTo E
Dim mRs As DAO.Recordset: Set mRs = CurrentDb.TableDefs(pNmtFm).OpenRecordset
If jj.Chk_Struct_Rs(mRs, Join(mAyFm, cComma)) Then ss.A 5: GoTo E
Dim mSel$: mSel = jj.ToStr_Am(mAm, " as ", "[]", "[]")
Dim mSql$: mSql = jj.Fmt_Str("Select {0} into {1} from {2}", mSel, jj.Rmv_SqBkt(pNmtTo), jj.Rmv_SqBkt(pNmtFm))
If jj.Run_Sql(mSql) Then ss.A 6: GoTo E
Exit Function
R: ss.R
E: Cv_Tbl2Tbl = True: ss.B cSub, cMod, "pNmtFm,pNmtTo,pLm", pNmtFm, pNmtTo, pLm
End Function
Function Cv_TypDAO_FmFldDcl(oTypDao As DAO.DataTypeEnum, oLen As Byte, pFldDcl$) As Boolean
Const cSub$ = "Cv_TypDAO_FmFldDcl"
On Error GoTo R
If Left(pFldDcl, 5) = "TEXT " Then
    oTypDao = dbText
    oLen = CByte(mID(pFldDcl, 6))
    Exit Function
End If
Select Case pFldDcl
Case "CURRENCY":     oTypDao = dbCurrency
Case "LONG", "AUTO": oTypDao = dbLong
Case "INT":          oTypDao = dbInteger
Case "BYTE":         oTypDao = dbByte
Case "DATE":         oTypDao = dbDate
Case "SINGLE":       oTypDao = dbSingle
Case "DOUBLE":       oTypDao = dbDouble
Case "MEMO":         oTypDao = dbMemo
Case "YESNO":        oTypDao = dbBoolean
Case Else: ss.A 1, "Invalid pFldDcl": GoTo E
End Select
Exit Function
R: ss.R
E: Cv_TypDAO_FmFldDcl = True: ss.B cSub, cMod, "pFldDcl", pFldDcl
End Function
Function Cv_Fld2Dcl$(pFld As DAO.Field)
Select Case pFld.Type
Case DAO.DataTypeEnum.dbChar _
   , DAO.DataTypeEnum.dbText: Cv_Fld2Dcl = "Text(" & pFld.Size & ")"
Case Else
                              Cv_Fld2Dcl = Cv_TypDAO_ToFldDcl(pFld.Type)
End Select
End Function
Function Cv_TypDAO_ToFldDcl$(pTypDAO As DAO.DataTypeEnum, Optional pLen As Byte)
Select Case pTypDAO
Case DAO.DataTypeEnum.dbBigInt _
   , DAO.DataTypeEnum.dbLong
                                    Cv_TypDAO_ToFldDcl = "Long"
Case DAO.DataTypeEnum.dbByte
                                    Cv_TypDAO_ToFldDcl = "Byte"
Case DAO.DataTypeEnum.dbCurrency _
   , DAO.DataTypeEnum.dbDecimal
                                    Cv_TypDAO_ToFldDcl = "Currency"
   
Case DAO.DataTypeEnum.dbDouble _
   , DAO.DataTypeEnum.dbFloat _
   , DAO.DataTypeEnum.dbNumeric
                                    Cv_TypDAO_ToFldDcl = "Double"
Case DAO.DataTypeEnum.dbInteger
                                    Cv_TypDAO_ToFldDcl = "Integer"
Case DAO.DataTypeEnum.dbSingle
                                    Cv_TypDAO_ToFldDcl = "Single"
Case DAO.DataTypeEnum.dbMemo
                                    Cv_TypDAO_ToFldDcl = "Memo"
Case DAO.DataTypeEnum.dbChar _
    , DAO.DataTypeEnum.dbText
                                    Cv_TypDAO_ToFldDcl = "Text " & pLen
Case DAO.DataTypeEnum.dbBoolean
                                    Cv_TypDAO_ToFldDcl = "YesNo"
Case DAO.DataTypeEnum.dbDate _
    , DAO.DataTypeEnum.dbTime _
    , DAO.DataTypeEnum.dbTimeStamp
                                    Cv_TypDAO_ToFldDcl = "Date"
Case Else
                                    Cv_TypDAO_ToFldDcl = "Unexpect TypDAO (" & pTypDAO & ")"
End Select
End Function
Function Cv_TypDAO2Sim(pTypDAO As DAO.DataTypeEnum) As eTypSim
Select Case pTypDAO
Case DAO.DataTypeEnum.dbBigInt _
    , DAO.DataTypeEnum.dbByte _
    , DAO.DataTypeEnum.dbCurrency _
    , DAO.DataTypeEnum.dbDecimal _
    , DAO.DataTypeEnum.dbDouble _
    , DAO.DataTypeEnum.dbFloat _
    , DAO.DataTypeEnum.dbInteger _
    , DAO.DataTypeEnum.dbLong _
    , DAO.DataTypeEnum.dbNumeric _
    , DAO.DataTypeEnum.dbSingle
                                    Cv_TypDAO2Sim = eTypSim_Num
Case DAO.DataTypeEnum.dbChar _
    , DAO.DataTypeEnum.dbMemo _
    , DAO.DataTypeEnum.dbText
                                    Cv_TypDAO2Sim = eTypSim_Str
Case DAO.DataTypeEnum.dbBoolean
                                    Cv_TypDAO2Sim = eTypSim_Bool
Case DAO.DataTypeEnum.dbDate _
    , DAO.DataTypeEnum.dbTime _
    , DAO.DataTypeEnum.dbTimeStamp
                                    Cv_TypDAO2Sim = eTypSim_Dte
Case Else
                                    Cv_TypDAO2Sim = eTypSim_Oth
End Select
End Function
Function Cv_TypSim2Chr$(pTyp As eTypSim)
Select Case pTyp
Case eTypSim_Bool: Cv_TypSim2Chr = "B"
Case eTypSim_Dte: Cv_TypSim2Chr = "D"
Case eTypSim_Num: Cv_TypSim2Chr = "N"
Case eTypSim_Oth: Cv_TypSim2Chr = "O"
Case eTypSim_Str: Cv_TypSim2Chr = "S"
Case Else: Cv_TypSim2Chr = "?"
End Select
End Function
Function Cv_V2Sim(pV) As eTypSim
Cv_V2Sim = Cv_VbVarTyp2Sim(VarType(pV))
End Function
Function Cv_VbVarTyp2Sim(pVbVarTyp As VbVarType) As eTypSim
Select Case pVbVarTyp
Case VBA.VbVarType.vbByte _
    , VBA.VbVarType.vbCurrency _
    , VBA.VbVarType.vbDecimal _
    , VBA.VbVarType.vbDouble _
    , VBA.VbVarType.vbInteger _
    , VBA.VbVarType.vbLong _
    , VBA.VbVarType.vbSingle
                                    Cv_VbVarTyp2Sim = eTypSim_Num
Case VBA.VbVarType.vbString
                                    Cv_VbVarTyp2Sim = eTypSim_Str
Case VBA.VbVarType.vbBoolean
                                    Cv_VbVarTyp2Sim = eTypSim_Bool
Case VBA.VbVarType.vbDate
                                    Cv_VbVarTyp2Sim = eTypSim_Dte
Case Else
                                    Cv_VbVarTyp2Sim = eTypSim_Oth
End Select
End Function
Function Cv_Vraw2Val(oVal, pVraw$, pTypSim As eTypSim) As Boolean
Const cSub$ = "Vraw2Val"
On Error GoTo R
Select Case pTypSim
Case eTypSim_Bool: oVal = CBool(pVraw)
Case eTypSim_Dte: oVal = CDate(pVraw): If oVal < #1/1/1990# Then ss.A 1, "Date is less <1990/1/1": GoTo E
Case eTypSim_Num: oVal = CDbl(pVraw)
Case eTypSim_Oth: ss.A 2, "TypSim(Other) is not handled", "pVraw", pVraw: GoTo E
Case eTypSim_Str: oVal = pVraw
Case Else: ss.A 3, "Unexpected TypSim", "pVraw,TypSim", pVraw, pTypSim: GoTo E
End Select
Exit Function
R: ss.R
E: Cv_Vraw2Val = True: ss.B cSub, cMod, "pVraw,pTypSim", pVraw, pTypSim
End Function
Function Cv_Dte2YYYYMM(pDte As Date) As tYYYYMM
Dim mYYYYMM As tYYYYMM
With mYYYYMM
    .YYYY = Year(pDte)
    .MM = Month(pDte)
End With
Cv_Dte2YYYYMM = mYYYYMM
End Function
Function Cv_YYYYMM2Nxt(pYYYYMM As tYYYYMM) As tYYYYMM
Dim mYYYYMM As tYYYYMM
With pYYYYMM
    If .MM = 12 Then
        mYYYYMM.MM = 1
        mYYYYMM.YYYY = .YYYY + 1
    Else
        mYYYYMM.MM = .MM + 1
        mYYYYMM.YYYY = .YYYY
    End If
End With
Cv_YYYYMM2Nxt = mYYYYMM
End Function
Function Cv_YYYYMM2Prv(pYYYYMM As tYYYYMM) As tYYYYMM
Dim mYYYYMM As tYYYYMM
With pYYYYMM
    If .MM = 1 Then
        mYYYYMM.MM = 12
        mYYYYMM.YYYY = .YYYY - 1
    Else
        mYYYYMM.MM = .MM - 1
        mYYYYMM.YYYY = .YYYY
    End If
End With
Cv_YYYYMM2Prv = mYYYYMM
End Function
Function Cv_YYYYMM2FirstDte(p As tYYYYMM) As Date
With p
    Cv_YYYYMM2FirstDte = CDate(.YYYY & "/" & .MM & "/1")
End With
End Function
Function Cv_YYYYMM2LasDte(p As tYYYYMM) As Date
With Cv_YYYYMM2Nxt(p)
    Cv_YYYYMM2LasDte = CDate(.YYYY & "/" & .MM & "/1") - 1
End With
End Function
