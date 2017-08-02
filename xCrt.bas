Attribute VB_Name = "xCrt"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xCrt"
Function Crt_Tbl_FmImpWs(pRge As Range, Optional pRithImp% = -1, Optional pIsDelta As Boolean = False, Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: create a table of name [>{mNmWs}] by NmFld row @ pRge(0,1) and import row {pRnoImp}, which contains valid TypFld for Crt_Tbl_FmLoFld.
Const cSub$ = "Crt_Tbl_FmImpWs"
On Error GoTo R
Dim mNmWs$: mNmWs = pRge.Parent.Name
Dim mNmt$
Dim mAmFld() As tMap, mAyCno() As Byte: If jj.Fnd_AyCnoImpFld(mAyCno, mAmFld, pRge, pRithImp) Then ss.A 1: GoTo E
Dim mLoFld$: mLoFld = ToStr_Am(mAmFld, " ")
If pIsDelta Then
    mNmt = ">>" & mNmWs
    If Not IsTbl(mNmt, pDb) Then
        If jj.Crt_Tbl_FmLoFld(mNmt, "Delta Long,TypDelta Text 1," & mLoFld, , , pDb) Then ss.A 2: GoTo E
        If jj.Crt_Tbl_FmLoFld(">#" & mNmWs, "TypDelta Text 1," & mLoFld, , , pDb) Then ss.A 2: GoTo E
    End If
    GoTo X
End If
If jj.Crt_Tbl_FmLoFld(">" & mNmWs, mLoFld, , , pDb) Then ss.A 2: GoTo E
GoTo X
R: ss.R
E: Crt_Tbl_FmImpWs = True: ss.B cSub, cMod, "mNmWs,pRge,pRithImp%", mNmWs, jj.ToStr_Rge(pRge), pRithImp%
X:
End Function
#If Tst Then
Function Crt_Tbl_FmImpWs_Tst() As Boolean
Dim mNmWs$, mRge As Range, mRnoImp&
mNmWs$ = "TblF"
If jj.Cpy_Fil("p:\AppDef_Meta\MetaDb.xls", "c:\tmp\aa.xls", True) Then Stop: GoTo E
Dim mWb As Workbook: If jj.Opn_Wb_R(mWb, "c:\tmp\aa.xls", , True) Then Stop: GoTo E
Dim mWs As Worksheet: Set mWs = mWb.Sheets(mNmWs)
Set mRge = mWs.Range("A5")
' mRnoImp = 1
If Crt_Tbl_FmImpWs(mRge) Then Stop: GoTo E
If jj.Opn_Tbl(">" & mNmWs, True) Then Stop: GoTo E
If Crt_Tbl_FmImpWs(mRge, , True) Then Stop: GoTo E
If jj.Opn_Tbl(">>" & mNmWs, True) Then Stop: GoTo E
GoTo X
E: Crt_Tbl_FmImpWs_Tst = True
X: jj.Cls_Wb mWb, False, True
End Function
#End If
'Function Crt_Tbl_FmFfnSml(pFfnSml$) As Boolean
''Aim: create a table (3 types of fields: Memo, Double, Date) by {pSml} of fmt: <Sml><Rec><Fld1 TypSim='yy'>..</Fld><Fld2>..</Fld2>...</Rec></Sml>.
''     yy can be Num, Dte, Str.  If omitted, Str is assumed.  Num create Double, Str creates Memo field in the output table.
''     The file name format is SmlXXX.yyyymmdd.hhmmss.xml, where [>XXX] will be the table name.
''     If the table exist, just delete the record and add from the content of Sml
'Const cSub$ = "Crt_Tbl_FmFfnSml"
'Dim mNmt0$ ' The table to be created
'Do
'    Dim mFn$: mFn = jj.Nam_FilNam(pFfnSml)
'    Dim mA1$, mA2$, mA3$, mA4$: If jj.Brk_Str_To4Seg(mA1, mA2, mA3, mA4, mFn, ".") Then ss.A 1, "pFfnSml must be in 4 segments sep by .": GoTo E
'    If Left(mA1, 3) <> "SML" Then ss.A 2, "pFfnSml must begin with SML": GoTo E
'    If mA4 <> "XML" Then ss.A 3, "pFfnSml must end with XML": GoTo E
'    mNmt0 = ">" & mID(mA1, 4)
'Loop Until True
'
'Dim mDocSml As MSXML2.DOMDocument60: If jj.Fnd_DocSml_ByFfn(mDocSml, pFfnSml) Then ss.A 1: GoTo E
'Dim iEleRec As MSXML2.IXMLDOMElement
''Find Tbl Struct: Assume first record contains all fields
'Dim mAnFld$(), mAyFldDcl$()
'    Dim NFld%
'    Set iEleRec = mDocSml.ChildNodes(0).ChildNodes(0)
'    Dim iEleFld As IXMLDOMElement
'    For Each iEleFld In iEleRec.ChildNodes
'        Dim mFldDcl$
'        If IsNull(iEleFld.getAttribute("FldDcl")) Then
'            mFldDcl = "Memo"
'        Else
'            mFldDcl = iEleFld.getAttribute("FldDcl").Value
'        End If
'        Select Case mFldDcl
'        Case "Memo", "Double", "Date"
'        Case Else
'            ss.A 2, "Unexpected NmTypSim in EleFld", , "mFldDcl,Ele Tag", mFldDcl, iEleFld.tagName: GoTo E
'        End Select
'        ReDim Preserve mAnFld(NFld), mAyFldDcl(NFld)
'        mAnFld(NFld) = iEleFld.tagName
'        mAyFldDcl(NFld) = mFldDcl
'        NFld = NFld + 1
'    Next
'
''Create Table
'If Not xIs.IsTbl(mNmt0) Then
'    Dim J%, mLoFld$: mLoFld = ""
'    For J = 0 To NFld - 1
'        mLoFld = jj.Add_Str(mLoFld, mAnFld(J) & " " & mAyFldDcl(J))
'    Next
'    If jj.Crt_Tbl_FmLoFld(mNmt0, mLoFld) Then ss.A 3: GoTo E
'End If
'
''Create rec
'For Each iEleRec In mDocSml.ChildNodes(0).ChildNodes
'    Dim mSql$
'        Dim mLv$, mLnFld$
'        mLv = "": mLnFld = ""
'        For Each iEleFld In iEleRec.ChildNodes
'            Dim mV$, mIdx%
'            If jj.Fnd_Idx(mIdx, mAnFld, iEleFld.tagName) Then ss.A 4: GoTo E
'            If mIdx < 0 Then ss.A 5: GoTo E
'            Select Case mAyFldDcl(mIdx)
'            Case "Memo": mV = jj.Q_S(iEleFld.Text)
'            Case "Double": mV = iEleFld.Text
'            Case "Date": mV = jj.Q_S(iEleFld.Text, "#")
'            Case Else: ss.A 6: GoTo E
'            End Select
'            mLnFld = jj.Add_Str(mLnFld, iEleFld.tagName)
'            mLv = jj.Add_Str(mLv, mV)
'        Next
'    mSql = jj.Fmt_Str("Insert into [{0}] ({1}) values ({2})", mNmt0, mLnFld, mLv)
'    If jj.Run_Sql(mSql) Then ss.A 4: GoTo E
'Next
'Exit Function
'R: ss.R
'E: Crt_Tbl_FmFfnSml = True: ss.B cSub, cMod, "pFfnSml", pFfnSml
'End Function
'#If Tst Then
'Function Crt_Tbl_FmFfnSml_Tst() As Boolean
'If jj.Exp_Str_ToFfn("<Sml><Rec><aa FldDcl='Num'>123</aa><bb>lskdjf</bb></Rec></Sml>", "c:\tmp\sml.xml", True) Then Stop: GoTo E
'If jj.Crt_Tbl_FmFfnSml("c:\tmp\SmlChgTbl.20080512.103549.xml") Then Stop: GoTo E
'DoCmd.OpenTable ">ChgTbl"
'Stop
'Exit Function
'E: Crt_Tbl_FmFfnSml_Tst = True
'End Function
'#End If
Function Crt_Idx_FmTbl(pNmt$) As Boolean
'Aim: Create Key for each record in {pNmt}: Fb,IsPk,IsNmUKey,IsUKey,IdxNo,LnFld
'If pIsPk    : Fb,NmTbl,     ,LnFld
'If pIsUKey  : Fb,NmTbl,IdxNo,LnFld
'If pIsNmUKey: Fb,NmTbl
Const cSub$ = "Crt_Idx_FmTbl"
Dim mSql$, mFbLas$, mRs As DAO.Recordset
mSql = "Select * from " & jj.Q_SqBkt(pNmt) & " order by Fb"
If jj.Opn_Rs(mRs, mSql) Then ss.A 1: GoTo E
With mRs
    While Not .EOF
        Dim mDb As DAO.Database
        If mFbLas <> !Fb Then
            jj.Cls_Db mDb
            mFbLas = !Fb
            If jj.Opn_Db_RW(mDb, mFbLas) Then ss.A 2: GoTo E
        End If
        Dim mNmTbl$: mNmTbl = "[$" & !NmTbl & "]"
        If !IsPk Then
            If jj.Crt_Pk(mNmTbl, !LnFld, mDb) Then ss.A 3: GoTo E
        ElseIf !IsNmUKey Then
            Dim mA$: mA = "Nm" & !NmTbl
            If jj.Crt_Idx(mNmTbl, mA, mA, True, mDb) Then ss.A 4: GoTo E
        Else
            Dim mNmIdx$: mNmIdx = IIf(!IsUniq, "U", "K") & Format(!IdxNo.Value, "00")
            If jj.Crt_Idx(mNmTbl, mNmIdx, !LnFld, !IsUniq, mDb) Then ss.A 4: GoTo E
        End If
        .MoveNext
    Wend
End With
Exit Function
E: Crt_Idx_FmTbl = True: ss.B cSub, cMod, "pNmt", pNmt
X:
    jj.Cls_Db mDb
    jj.Cls_Rs mRs
End Function
Function Crt_Rge_ExtNCol(oRge As Range, pRge As Range, Optional pNCol As Byte = 1) As Boolean
'Aim: assume pRge is a rectangle. Create oRge which by extending pRge by pNCol more column
Const cSub$ = "Crt_Rge_ExtNCol"
On Error GoTo R
Dim mR1&, mC1 As Byte
Dim mR2&, mC2 As Byte
With pRge
    mR1 = .Row
    mC1 = .Column
    mR2 = mR1 + .Rows.Count - 1
    mC2 = mC1 + .Columns.Count - 1
End With
Set oRge = jj.Crt_Rge_Fm2Pts(pRge.Parent, mR1, mC1, mR2, mC2 + pNCol)
Exit Function
R: ss.R
E: Crt_Rge_ExtNCol = True: ss.B cSub, cMod, "pRge,pNCol", jj.ToStr_Rge(pRge), pNCol
End Function
Function Crt_Tbl_ParChd(pNmtTar$, pNmtSrc$, pPar$, pChd$) As Boolean
'Aim: Build pNmtTar of structure: Sno, Par, Chd, Lvl, from {pNmtSrc} & {pNmtNm}
'     Assume Struct: pNmtTar: {pPar}, {pChd}
Const cSub$ = "Crt_Tbl_ParChd"
Dim mNmtSrc$: mNmtSrc = jj.Rmv_SqBkt(pNmtSrc)
If jj.Chk_Struct_Tbl_SubSet(mNmtSrc, pPar & "," & pChd) Then ss.A 1: GoTo E
Dim mAyRoot&(): If jj.Fnd_AyRoot(mAyRoot, pNmtSrc, pPar, pChd) Then ss.A 3: GoTo E
Dim mNmtTar$: mNmtTar = jj.Rmv_SqBkt(pNmtTar)
If jj.Crt_Tbl_FmLoFld(mNmtTar, "Sno Long, Par Long, Chd Long, Lvl Byte", 1) Then ss.A 4: GoTo E
Dim mRsTar As DAO.Recordset: If jj.Opn_Rs(mRsTar, "Select * from [" & mNmtTar & "]") Then ss.A 5: GoTo E
Dim mSno&, mLvl As Byte: mSno = 0: mLvl = 0
Dim J%
Dim mAyPth&(), N%: N = jj.Siz_Ay(mAyRoot)
For J = 0 To N - 1
    If J Mod 50 = 0 Then jj.Shw_Sts J & "(" & N & ") ..."
    Crt_Tbl_ParChd_OneRec mRsTar, 0, mAyRoot(J), mLvl
    If Crt_Tbl_ParChd_OneRoot(mAyRoot(J), mAyPth, mLvl, mRsTar, mNmtSrc, pPar, pChd) Then ss.A 6: GoTo E
Next
GoTo X
E: Crt_Tbl_ParChd = True: ss.B cSub, cMod, "pNmtTar,pNmtSrc,pPar,pChd", pNmtTar, pNmtSrc, pPar, pChd
X: jj.Cls_Rs mRsTar
End Function
#If Tst Then
Function Crt_Tbl_ParChd_Tst() As Boolean
'If jj.Crt_Tbl_FmLnkLnt("p:\workingdir\MetaDb.mdb", "$Tbl,$TblR") Then Stop: GoTo E
Dim mFx$, mWb1 As Workbook, mWb2 As Workbook, mWs As Worksheet
If True Then
    mFx = "c:\tmp\aa.xls"
    If True Then
        If jj.Crt_Tbl_FmLnkLnt("P:\WorkingDir\MetaAll.mdb", "$Tbl,$TblR") Then Stop: GoTo E
        If jj.Run_Qry("qryTstCrtTblParChd") Then Stop: GoTo E
        If jj.Exp_SetNmtq2Xls("[#]Lst", mFx, True) Then Stop: GoTo E
    End If
    If jj.Opn_Wb_RW(mWb1, mFx) Then Stop: GoTo E
    Set mWs = mWb1.Sheets(1)
    If jj.Fmt_WsOL_ByCol(mWs.Range("A2"), 5, 6) Then Stop: GoTo E
    mWb1.Save
    mWb1.Application.Visible = True
    Stop
End If
If True Then
    mFx = "c:\tmp\bb.xls"
    If jj.Crt_Tbl_ParChd("#Tmp", "$TblR", "TblTo", "Tbl") Then Stop: GoTo E
    
    If jj.Run_Sql("Alter table [#Tmp] Add NmPar Text(50), L Long, NmChd Text(50)") Then Stop: GoTo E
    If jj.Run_Sql("Update [#Tmp] m inner join [$Tbl] s" & _
        " On m.Par=s.Tbl" & _
        " Set m.NmPar=s.NmTbl" & _
        " Where Par<>0") Then Stop: GoTo E
    If jj.Run_Sql("Update [#Tmp] set NmPar='Root' where Par=0") Then Stop: GoTo E
    If jj.Run_Sql("Update [#Tmp] set L=Lvl+1") Then Stop: GoTo E
    If jj.Run_Sql("Alter Table [#Tmp] Drop Column Lvl") Then Stop: GoTo E
    If jj.Run_Sql("Update [#Tmp] m inner join [$Tbl] s" & _
        " On m.Chd=s.Tbl" & _
        " Set m.NmChd=s.NmTbl") Then Stop: GoTo E
    
    If jj.Exp_SetNmtq2Xls("[#]Tmp", mFx, True) Then Stop: GoTo E
    If jj.Opn_Wb_RW(mWb2, mFx) Then Stop: GoTo E
    Set mWs = mWb2.Sheets(1)
    If jj.Fmt_WsOL_ByCol(mWs.Range("A2"), 5, 6) Then Stop: GoTo E
    mWb2.Save
End If
mWs.Application.Visible = True
Stop
GoTo X
Exit Function
E: Crt_Tbl_ParChd_Tst = True
X:
    jj.Cls_Wb mWb1, False, True
    jj.Cls_Wb mWb2, False, True
End Function
#End If
Private Function Crt_Tbl_ParChd_OneRoot(ByVal pRoot&, oAyPth&(), oLvl As Byte, pRsTar As DAO.Recordset, pNmtSrc$, pPar$, pChd$) As Boolean
'Aim: Recursively write records to {pRsTar}.  Each root one extra write.
'     pRsTar: Sno, Par, Chd, Lvl
'     Assume pNmtSrc has no []
Const cSub$ = "Crt_Tbl_ParChd_OneRoot"
On Error GoTo R
oLvl = oLvl + 1

Dim mSql$: mSql = jj.Fmt_Str("Select {0} from [{1}] where {2}={3} order by {0}", pChd, pNmtSrc, pPar, pRoot)
Dim mAyId&(): If jj.Fnd_AyVFmSql(mAyId, mSql) Then ss.A 2: GoTo E

Dim J%
For J = 0 To jj.Siz_Ay(mAyId) - 1
    Crt_Tbl_ParChd_OneRec pRsTar, pRoot, mAyId(J), oLvl
    
    Dim mIdx%: If jj.Fnd_IdxLng(mIdx, oAyPth, mAyId(J)) Then ss.A 3: GoTo E
    If mIdx < 0 Then
        Dim N%
        N = jj.Siz_Ay(oAyPth)
        ReDim Preserve oAyPth(N): oAyPth(N) = mAyId(J)
        If Crt_Tbl_ParChd_OneRoot(mAyId(J), oAyPth, oLvl, pRsTar, pNmtSrc, pPar, pChd) Then ss.A 4: GoTo E
        If N = 0 Then
            jj.Clr_AyLng oAyPth
        Else
            ReDim Preserve oAyPth(N - 1)
        End If
    End If
Next

oLvl = oLvl - 1
Exit Function
R: ss.R
E: Crt_Tbl_ParChd_OneRoot = True: ss.B cSub, cMod, "pRoot,oLvl,pRsTar,pNmtSrc,pPar,pChd", pRoot, oLvl, jj.ToStr_Rs_NmFld(pRsTar), pNmtSrc, pPar, pChd
End Function
Private Function Crt_Tbl_ParChd_OneRec(pRsTar, pPar&, pChd&, pLvl As Byte) As Boolean
'     pRsTar: Sno, Par, Chd, Lvl
With pRsTar
    .AddNew
    !Par = pPar
    !Chd = pChd
    !Lvl = pLvl
    .Update
End With
End Function
Function Crt_Fx(pFx$, Optional pOvrWrt As Boolean = False, Optional pNmWs$ = "ToBeDelete") As Boolean
Const cSub$ = "Crt_Fx"
Dim mWb As Workbook: If jj.Crt_Wb(mWb, pFx, pOvrWrt, pNmWs) Then ss.A 1: GoTo E
If jj.Cls_Wb(mWb, True) Then ss.A 2: GoTo E
Exit Function
E: Crt_Fx = True: ss.B cSub, cMod, "pFx,pOvrWrt,pNmWs", pFx, pOvrWrt, pNmWs
End Function
#If Tst Then
Function Crt_Fx_Tst() As Boolean
If jj.Crt_Fx("c:\tmp\aa.xls", True, "Sheet1") Then Stop
End Function
#End If
Function Crt_DQry() As d_Qry
Dim mDQry As New d_Qry: Set Crt_DQry = mDQry
End Function
Function Crt_Fld_FmRsTblF(oFld As DAO.Field, pRsTblF As DAO.Recordset) As Boolean
'     #TblF: NmFld,TypDao,FldLen,FmtTxt,IsReq,IsAlwZerLen,DftVal,VdtTxt,VdtRul
Const cSub$ = "Crt_Fld_FmRsTblF"
On Error GoTo R
Dim mNmFld$, mTyp As DAO.DataTypeEnum, mSiz As Byte, mIsAuto As Boolean, mIsReq As Boolean, mAlwZerLen As Boolean, mDftVal$, mFmtTxt$, mVdtTxt$, mVdtRul$
With pRsTblF
    mNmFld = !NmFld
    mTyp = !TypDao
    mSiz = Nz(!FldLen, 0)
    mFmtTxt = Nz(!FmtTxt, "")
    mIsReq = !IsReq
    mDftVal = Nz(!DftVal.Value, "")
    Select Case mTyp
    Case DAO.DataTypeEnum.dbText, DAO.DataTypeEnum.dbMemo
        mAlwZerLen = !IsAlwZerLen
    Case Else
        mAlwZerLen = False
    End Select
    mVdtTxt = Nz(!VdtTxt, "")
    mVdtRul = Nz(!VdtRul, "")
End With
If jj.Crt_Fld(oFld, mNmFld, mTyp, mSiz, mIsAuto, mAlwZerLen, mIsReq, mDftVal, mFmtTxt, mVdtTxt, mVdtRul) Then ss.A 3: GoTo E
Exit Function
R: ss.R
E: Crt_Fld_FmRsTblF = True: ss.B cSub, cMod, "pRsTblF", jj.ToStr_Rs(pRsTblF)
End Function
#If Tst Then
Function Crt_Fld_FmRsTblF_Tst() As Boolean

End Function
#End If
Function Crt_Fld(oFld As DAO.Field, pNmFld$, pTyp As DAO.DataTypeEnum _
    , Optional pSiz As Byte = 0 _
    , Optional pIsAuto As Boolean = False _
    , Optional pAlwZerLen As Boolean = False _
    , Optional pIsReq As Boolean = False _
    , Optional pDftVal$ _
    , Optional pFmtTxt$ = "" _
    , Optional pVdtTxt$ = "" _
    , Optional pVdtRul$ = "" _
    ) As Boolean
Const cSub$ = "Crt_Fld"
Set oFld = New DAO.Field
On Error GoTo R
With oFld
    .Name = jj.Rmv_SqBkt(pNmFld)
    If pTyp = 0 Then ss.A 1, "pTyp cannot be zero"
    .Type = pTyp
    'If pTyp = dbMemo Then Stop
    If pSiz > 0 Then .Size = pSiz
    If .AllowZeroLength <> pAlwZerLen Then .AllowZeroLength = pAlwZerLen
    If pDftVal <> "" Then
        If pTyp = dbText Then
            .DefaultValue = jj.Q_S(pDftVal, """")
        Else
            .DefaultValue = pDftVal
        End If
    End If
    .Required = pIsReq
    If pFmtTxt <> "" Then oFld.Properties.Append oFld.CreateProperty("Format", DAO.DataTypeEnum.dbText, pFmtTxt)
    If pVdtTxt <> "" Then .ValidationText = pVdtTxt
    If pVdtRul <> "" Then .ValidationRule = pVdtRul
    If pIsAuto Then .Attributes = DAO.FieldAttributeEnum.dbAutoIncrField
End With
Exit Function
R: ss.R
E: Crt_Fld = True: ss.B cSub, cMod, "pNmFld,pTyp,pSiz,pIsAuto,pAlwZerLen,pIsReq,pDftVal,pFmtTxt,pVdtTxt,pVdtTxt,pVdtRul", pNmFld$, jj.ToStr_TypDta(pTyp), pSiz, pIsAuto, pAlwZerLen, pIsReq, pDftVal, pFmtTxt, pVdtTxt, pVdtTxt, pVdtRul
End Function
Function Crt_Tbl_ForTxtSpec(Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: Create 2 tables (MSysIMEXSpecs & MSysIMEXColumns) in {pDb}
Const cSub$ = "Crt_Tbl_ForTxtSpec"
If Not jj.IsTbl("MSysIMEXSpecs", pDb) Then
    If jj.Crt_Tbl_FmLoFld("MSysIMEXSpecs", _
        "SpecName Text 64" & _
        ", SpecId Auto" & _
        ", DateDelim Text 2" & _
        ", DateFourDigitYear YesNo" & _
        ", DateLeadingZeros YesNo" & _
        ", DecimalPoint Text 2" & _
        ", DateOrder Int" & _
        ", FieldSeparator Text 2" & _
        ", FileType Int" & _
        ", SpecType Byte" & _
        ", StartRow Long" & _
        ", TextDelim Text 2" & _
        ", TimeDelim Text 2" _
        , 1, 2, pDb) Then ss.A 1: GoTo E
End If
If Not jj.IsTbl("MSysIMEXColumns", pDb) Then
    If jj.Crt_Tbl_FmLoFld("MSysIMEXColumns", _
        "SpecId Long" & _
        ", FieldName Text 64" & _
        ", Attributes Long" & _
        ", DataType Int" & _
        ", IndexType Byte" & _
        ", SkipColumn YesNo" & _
        ", Start Int" & _
        ", Width Int" _
        , 2, 2, pDb) Then ss.A 2: GoTo E
    Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
    mDb.Execute "Create index Index1 on MSysIMEXColumns (SpecId,Start)"
End If
Exit Function
R: ss.R
E: Crt_Tbl_ForTxtSpec = True: ss.B cSub, cMod, "pDb", jj.ToStr_Db(pDb)
'     Format of pLoFld is xxx Text 10,....
'     Note: xxx may be in xx^xx format.  ^ means for space
'       TEXT,CURRENCY,LONG,INT,BYTE,DATE,SINGLE,DOUBLE,MEMO,YESNO
End Function
Function Crt_Tbl_ForTxtSpec_Tst() As Boolean
Dim mDb As DAO.Database: If jj.Crt_Db(mDb, "c:\aa.mdb", True) Then Stop
If jj.Crt_Tbl_ForTxtSpec(mDb) Then Stop
Stop
End Function
Function Crt_Tbl_FmLnkTxt(pNmt$, pFt$, pNmSpec$, Optional pInDb As DAO.Database = Nothing) As Boolean
'Aim: Create {pNmt} in {pInDb} by linking {pFt} using {pNmSpec}.
Const cSub$ = "Crt_Tbl_FmLnkTxt"
If VBA.Dir(pFt) = "" Then ss.A 1, "Given txt file not found": GoTo E
'Text;DSN=A1 Link Specification;FMT=Fixed;HDR=NO;IMEX=2;CharacterSet=20127;DATABASE=C:\;TABLE=a1#txt
Dim mDir$: mDir = jj.Fct.Nam_DirNam(pFt)
Dim mNmtSrc$:  mNmtSrc = jj.Fct.Nam_FilNam(pFt)
Dim mFn$: mFn = Replace(mNmtSrc, ".", "#")
Dim mCnn$: mCnn = jj.Fmt_Str("Text;DSN={0};FMT=Fixed;HDR=NO;IMEX=2;CharacterSet=20127;DATABASE={1};TABLE={2}", pNmSpec, mDir, mFn)
Crt_Tbl_FmLnkTxt = jj.Crt_Tbl_FmLnk(pNmt, mNmtSrc, mCnn, pInDb)
Exit Function
R: ss.R
E: Crt_Tbl_FmLnkTxt = True: ss.B cSub, cMod, "pNmt$, pFt$, pNmSpec$, pInDb", pNmt$, pFt$, pNmSpec$, jj.ToStr_Db(pInDb)
End Function
#If Tst Then
Function Crt_Tbl_FmLnkTxt_Tst() As Boolean
Dim mNmt$: mNmt = "A1"
Dim mFb$: mFb = "c:\aa.mdb"
Dim mFt$: mFt = "c:\a1.txt"
Dim mNmSpec$: mNmSpec = "A1"
Dim mTxtSpec$: mTxtSpec = "I=Int3, AA=Txt10, B=Txt2, C=Txt3"
Dim mDb As DAO.Database:: If jj.Crt_Db(mDb, mFb, True) Then Stop
If jj.Dlt_Tbl(mNmt, mDb) Then Stop
If jj.Dlt_TxtSpec(mNmSpec, mDb) Then Stop
If jj.Crt_TxtSpec_Fix(mNmSpec, mTxtSpec, mDb) Then Stop
If jj.Dlt_Fil(mFt) Then Stop
Open mFt For Output As #1
Close #1
If jj.Crt_Tbl_FmLnkTxt(mNmt, mFt, mNmSpec, mDb) Then Stop
End Function
#End If
Function Crt_TxtSpec_Delimi(pNmSpec$, pAmFld() As tMap, Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: Delete and Add one record to MSysIMEXSpecs & N records to MSysIMEXColumns to create a "text" file link spec
'     MSysIMEXSpecs  : DateDelim,DateFourDigitYear,DateLeadingZeros,DateOrder,DecimalPoint,FieldSeparator,FileType,SpecID,SpecName,SpecType,StartRow,TextDelim,TimeDelim
'     MSysIMEXColumns: Attributes,DataType,FieldName,IndexType,SkipColumn,SpecID,Start,Width
'     TxtSpec is in Am format <NmFld>=<Spec>;^^^
'     Note: <Spec>:TEXT NN,CURRENCY,LONG,INT,BYTE,DATE,SINGLE,DOUBLE,MEMO,YESNO
'           YesNo    always len=1
'           DateTime always len=8 + 1 + 6
'Hdr
'DateDelim   /
'DateFourDigitYear True
'DateLeadingZeros False
'DateOrder 5
'DecimalPoint    .
'FieldSeparator  ,
'FileType -536
'SpecID 3
'SpecName aa
'SpecType 1
'StartRow 1
'TextDelim ""
'TimeDelim:
'Det
'Attributes  0   0
'DataType    10  10
'FieldName   Obj NmObj
'IndexType   0   0
'SkipColumn  FALSE   FALSE
'SpecID  3   3
'Start   1   5
'Width   4   7

Const cSub$ = "Crt_TxtSpec_Delimi"
If jj.Dlt_TxtSpec(pNmSpec, pDb) Then ss.A 1: GoTo E 'Create the Spec Tables if not exist

'Create one record in MSysIMEXSpecs
Dim mSql$: mSql = jj.Fmt_Str( _
"Insert into MSysIMEXSpecs (DateDelim,DateFourDigitYear,DateLeadingZeros,DateOrder,DecimalPoint,FieldSeparator,FileType,SpecName,SpecType,StartRow,TextDelim,TimeDelim) values " & _
                          "('/'      ,True             ,Yes             ,5        ,'.'         ,','           ,-536    ,'{0}'   ,1       ,1       ,'""'     ,':')", pNmSpec)
If jj.Run_Sql_ByDbExec(mSql, pDb) Then ss.A 2: GoTo E

'Get SpecId by SpecName
Dim mSpecId&: If jj.Fnd_ValFmSql(mSpecId, "Select SpecId from MSysIMEXSpecs where SpecName='" & pNmSpec & cQSng, pDb) Then ss.A 1: GoTo E

'Attributes
'    DataType
'        FieldName   IndexType
'                        SkipColumn
'                            SpecID
'                                Start
'                                    Width
'0   3   INT         0   0   6   1   3
'0   8   DATETIME    0   0   6   4   15
'0   5   CUR         0   0   6   19  10
'0   12  MEMO        0   0   6   29  10
'0   4   LONG        0   0   6   39  10
'0   2   BYTE        0   0   6   49  3
'0   1   YESNO       0   0   6   52  10
'0   7   DOUBLE      0   0   6   62  10
'0   10  TEXT        0   0   6   72  10
'0   6   SINGLE      0   0   6   82  10
'TEXT NN,CURRENCY,LONG,INT,BYTE,DATE,SINGLE,DOUBLE,MEMO,YESNO
'Create N records to MSysIMEXColumns
Dim J%, mStart%, mWidth%
mStart = 1: mWidth = 0
For J = 0 To jj.Siz_Am(pAmFld) - 1
    With pAmFld(J)
        Dim mDtaTyp As Byte
        Do
            Select Case .F2
            Case "YesNo":    mDtaTyp = 1: mWidth = 1
            Case "Date": mDtaTyp = 8: mWidth = 8 + 1 + 6
            Case Else
                Dim mA$, mP%: mP = InStr(.F2, " ")
                If mP > 0 Then mA = Left(.F2, mP - 1) Else mA = .F2
                Select Case Trim(mA)
                Case "INT": mDtaTyp = 3:  mWidth = 1
                Case "CURRENCY": mDtaTyp = 5:  mWidth = 1
                Case "MEMO":     mDtaTyp = 12: mWidth = 1
                Case "LONG":     mDtaTyp = 4:  mWidth = 1
                Case "BYTE":     mDtaTyp = 2:  mWidth = 1
                Case "DOUBLE":   mDtaTyp = 7:  mWidth = 1
                Case "TEXT":     mDtaTyp = 10: mWidth = 1
                Case "SINGLE":   mDtaTyp = 6:  mWidth = 1
                Case Else
                    ss.A 4, "Invalid TypFld", eRunTimErr, "NmFld,FldSpec,Valid Spec", .F1, .F2, "TEXT NN,CURRENCY,LONG,INT,BYTE,DATE,SINGLE,DOUBLE,MEMO,YESNO": GoTo E
                End Select
            End Select
        Loop Until True
        
        mSql = jj.Fmt_Str( _
        "Insert into MSysIMEXColumns (Attributes,DataType,FieldName,IndexType,SkipColumn,SpecID,Start,Width) values " & _
                                    "(0         ,{0}     ,'{1}'    ,0        ,0         ,{2}   ,{3}  ,{4})", _
                                    mDtaTyp, .F1, mSpecId, mStart, mWidth)
    End With
    mStart = mStart + mWidth
    If jj.Run_Sql_ByDbExec(mSql, pDb) Then ss.A 5: GoTo E
Next
Exit Function
R: ss.R
E: Crt_TxtSpec_Delimi = True: ss.B cSub, cMod, "pNmSpec,pAmFld,pDb", pNmSpec, jj.ToStr_Am(pAmFld), jj.ToStr_Db(pDb)
End Function
#If Tst Then
Function Crt_TxtSpec_Delimi_Tst() As Boolean
Const cSub$ = "Crt_TxtSpec_Delimi_Tst"
Dim mWb As Workbook: If jj.Opn_Wb_R(mWb, "P:\AppDef_Meta\MetaDb.xls") Then ss.A 1: GoTo E
Dim mAnWs$(): If jj.Fnd_AnWs_BySetWs(mAnWs, mWb) Then ss.A 2: GoTo E
Dim J%, N%: N = jj.Siz_Ay(mAnWs)
Dim mXls As Excel.Application: Set mXls = mWb.Application: mXls.DisplayAlerts = False
ReDim mAyFfn$(N - 1)
For J = 0 To N - 1
    Dim mWs As Worksheet: Set mWs = mWb.Sheets(mAnWs(J))
    Dim mAmFld() As tMap, mAyCno() As Byte: If jj.Fnd_AyCnoImpFld(mAyCno, mAmFld, mWs.Range("A5")) Then ss.A 1: GoTo E

    If jj.Clr_ImpWs(mWs.Range("A5")) Then ss.A 1: GoTo E
    'Save to Csv
    mAyFfn(J) = "c:\tmp\Exp_Ws2Tbl" & jj.Fct.TimStmp & "_" & mWs.Name & ".csv"
    mWs.SaveAs mAyFfn(J), Excel.XlFileFormat.xlCSVWindows
    If jj.Crt_TxtSpec_Delimi(">" & mAnWs(J), mAmFld) Then ss.A 1: GoTo E
Next
jj.Cls_Wb mWb, False, True
If jj.Dlt_Tbl_ByPfx(">") Then ss.A 2: GoTo E
For J = 0 To N - 1
    DoCmd.TransferText acImportDelim, ">" & mAnWs(J), ">" & mAnWs(J), mAyFfn(J), True
Next
GoTo X
E: Crt_TxtSpec_Delimi_Tst = True: ss.B cSub, cMod
X: mXls.DisplayAlerts = True
   jj.Cls_Wb mWb, False, True
End Function
#End If
Function Crt_TxtSpec_Fix(pNmSpec$, pLmTxtSpec$, Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: Create (over-write}a Fixed len txt spec {pNmSpec} in {pDb} by {pLmTxtSpec}
'     Txt Spec are 2 tables definition: Delete and Add one record to MSysIMEXSpecs & N records to MSysIMEXColumns to create a "text" file link spec
'     MSysIMEXSpecs  : DateDelim,DateFourDigitYear,DateLeadingZeros,DateOrder,DecimalPoint,FieldSeparator,FileType,SpecID,SpecName,SpecType,StartRow,TextDelim,TimeDelim
'     MSysIMEXColumns: Attributes,DataType,FieldName,IndexType,SkipColumn,SpecID,Start,Width
'     TxtSpec is in Lm format <NmFld>=<Spec>;
'       <Spec>=Txt<n> Byt<n> Int<n> Sng<n> Dbl<n> Cur<n> Mem<n> YesNo DateTime
'           YesNo    always len=1
'           DateTime always len=8 + 1 + 6
Const cSub$ = "Crt_TxtSpec_Fix"
If jj.Dlt_TxtSpec(pNmSpec, pDb) Then ss.A 1: GoTo E

'Break pLmTxtSpec
Dim mAm() As tMap: mAm = Get_Am_ByLm(pLmTxtSpec)

'Create one record in MSysIMEXSpecs
Dim mSql$: mSql = jj.Fmt_Str( _
"Insert into MSysIMEXSpecs (DateDelim,DateFourDigitYear,DateLeadingZeros,DateOrder,DecimalPoint,FieldSeparator,FileType,SpecName,SpecType,StartRow,TextDelim,TimeDelim) values " & _
                          "(''       ,True             ,Yes             ,5        ,'.'         ,','           ,20127   ,'{0}'   ,2       ,0       ,''       ,'')", pNmSpec)
If jj.Run_Sql(mSql) Then ss.A 2: GoTo E

'Get SpecId by SpecName
Dim mSpecId&: If jj.Fnd_ValFmSql(mSpecId, "Select SpecId from MSysIMEXSpecs where SpecName='" & pNmSpec & cQSng) Then ss.A 1: GoTo E

'Attributes
'    DataType
'        FieldName   IndexType
'                        SkipColumn
'                            SpecID
'                                Start
'                                    Width
'0   3   INT         0   0   6   1   3
'0   8   DATETIME    0   0   6   4   15
'0   5   CUR         0   0   6   19  10
'0   12  MEMO        0   0   6   29  10
'0   4   LONG        0   0   6   39  10
'0   2   BYTE        0   0   6   49  3
'0   1   YESNO       0   0   6   52  10
'0   7   DOUBLE      0   0   6   62  10
'0   10  TEXT        0   0   6   72  10
'0   6   SINGLE      0   0   6   82  10
'Txt<n> Byt<n> Int<n> Sng<n> Dbl<n> Cur<n> Mem<n> YesNo DateTime

'Create N records to MSysIMEXColumns
Dim J%, mStart%, mWidth%
mStart = 1: mWidth = 0
For J = 0 To jj.Siz_Am(mAm) - 1
    With mAm(J)
        Dim mDtaTyp As Byte
        Do
            Select Case .F2
            Case "YesNo":    mDtaTyp = 1: mWidth = 1
            Case "DateTime": mDtaTyp = 8: mWidth = 8 + 1 + 6
            Case Else
                Dim mA$: mA = Left(.F2, 3)
                If Len(.F2) <= 3 Then ss.A 3, "Invalid data type", eRunTimErr, "NmFld,FldSpec,Valid Spec", .F1, .F2, "Txt<n> Byt<n> Int<n> Sng<n> Dbl<n> Cur<n> Mem<n> YesNo DateTime": GoTo E
                Select Case mA
                Case "INT": mDtaTyp = 3:  mWidth = mID(.F2, 4)
                Case "CUR": mDtaTyp = 5:  mWidth = mID(.F2, 4)
                Case "MEM": mDtaTyp = 12: mWidth = mID(.F2, 4)
                Case "LNG": mDtaTyp = 4:  mWidth = mID(.F2, 4)
                Case "BYT": mDtaTyp = 2:  mWidth = mID(.F2, 4)
                Case "DBL": mDtaTyp = 7:  mWidth = mID(.F2, 4)
                Case "TXT": mDtaTyp = 10: mWidth = mID(.F2, 4)
                Case "SNG": mDtaTyp = 6:  mWidth = mID(.F2, 4)
                Case Else
                    ss.A 4, "Invalid data type", eRunTimErr, "NmFld,FldSpec,Valid Spec", .F1, .F2, "Txt<n> Byt<n> Int<n> Sng<n> Dbl<n> Cur<n> Mem<n> YesNo DateTime": GoTo E
                End Select
            End Select
        Loop Until True
        
        mSql = jj.Fmt_Str( _
        "Insert into MSysIMEXColumns (Attributes,DataType,FieldName,IndexType,SkipColumn,SpecID,Start,Width) values " & _
                                    "(0         ,{0}     ,'{1}'    ,0        ,0         ,{2}   ,{3}  ,{4})", _
                                    mDtaTyp, .F1, mSpecId, mStart, mWidth)
    End With
    mStart = mStart + mWidth
    If jj.Run_Sql(mSql) Then ss.A 5: GoTo E
Next
Exit Function
R: ss.R
E: Crt_TxtSpec_Fix = True: ss.B cSub, cMod, "pNmSpec,pLmTxtSpec,pDb", pNmSpec, pLmTxtSpec, jj.ToStr_Db(pDb)
End Function
Function Crt_TxtSpec_Fix_Tst() As Boolean
If jj.Crt_TxtSpec_Fix("A2Test", "I=Int3, A=Txt1, B=Txt2, C=Txt3") Then Stop: GoTo E
Stop
Dim mF As Byte: If jj.Opn_Fil_ForOutput(mF, "c:\tmp\aa.txt", True) Then Stop: GoTo E
Print #mF, "123XAA 22"
Print #mF, "12 YAB  2"
Print #mF, "1  ZAB   "
Print #mF, "123 AB222"
Close #mF
DoCmd.TransferText acImportFixed, "A2Test", "#Tmp", "c:\tmp\aa.txt", False
DoCmd.OpenTable "#Tmp"
Stop
GoTo X
E: Crt_TxtSpec_Fix_Tst = True
X: Close mF
End Function
Function Crt_Tbl_FmAySql_ByDsn(pNmtTar$, pAyDsn$(), pAySelSql$() _
        , Optional pFbTar$ = "" _
        ) As Boolean
Const cSub$ = "Crt_Tbl_FmAySql_ByDsn"
'Aim: Create {pNmtTar} in {pFbTar} by downloading data from {pAyDsn} by {pAySelSql}
'     {pAySelSql} should select same structure may from different source {pAyDsn}
'----

Dim mAnQ1$()
Do ' Bld Qs
    Do ' Bld Q0
    
    Loop Until True
    
    Do ' Bld Q1
    Loop Until True
    
    Do ' Bld Q2
    Loop Until True
Loop Until True

Do ' Run_Qs
    If jj.Run_Qry_ByAnq(mAnQ1) Then Stop
Loop Until True
Exit Function
E: Crt_Tbl_FmAySql_ByDsn = True
End Function
''Notes:
''     Q1    # of queries in Q1 will be same as # of element in pAy*.
''           Running all Q1 will download data to [mNmtTar]
''           The first query of Q1 is Create Table query, while the rest are Append Table query
''     [ClearUp] If {pRun} & not {gOdbcDbg} then delete the [Q1 & Q2] & [Dtf import tables], so that only tmp{mNmqsns}_{pNmDl} will left
''           [Q2]: Select * from {[mNmtTar_DTF]}
''           [Q1]: Standard,   Const cSql_Crt$ = "Select {0} into {1} from {2}"
''                             Const cSql_App$ = "Insert into {1} Select {0} from {2}"
''
''Assume: the pDsn is used to create the library, which means there will be no library in the ODBC query.
'''    Name                                  Sql
'''Q0: qryOdbc{pNmQs}_{pMajNo}_0_{pNmDl}     Select * from tmp{pNmQs}_{pNmDl}
'''Q1: qryOdbc{pNmQs}_{pMajNo}_1_{n}Fm_Below Select * into tmp{pNmQs}_{pNmDl} from {qry3n}
'''Q2: qryOdbc{pNmQs}_{pMajNo}_2_{n}{xNmDl}  {pAySelSql()}
'Dim N%: N = jj.Siz_Ay(pAyDsn)
'If jj.Siz_Ay(pAySelSql) <> N Then ss.A 1, "Siz of pAyDsn and pAySelSql must be the same", "SizOf pAyDsn, SizOf pAySelSql", N, jj.Siz_Ay(pAySelSql): Goto E
'
''Build Query 0
'Dim mNmq0$: mNmq0 = jj.Fmt_Str("qryOdbc_0_Crt_{0}", pNmtTar)
'Dim mInFbOupTbl$: If pFbTar <> "" Then mInFbOupTbl = " in '" & pFbTar & cQSng
'Dim mSql$: mSql = jj.Fmt_Str("Select * from {0}", pNmtTar)
'If jj.Crt_Qry(mNmq0, mSql) Then ss.A 2:Goto E
'
'Dim J%: For J = 0 To N - 1
'    Dim A$: A = IIf(N = 1, "", J)
'     'Build Query 2
'    Dim mNmq1$: mNmq1 = jj.Fmt_Str("qryOdbc_1_{0}Fm_Below", A)
'    Dim mNmq2$: mNmq2 = jj.Fmt_Str("qryOdbc_2_{0}", A)
'    Dim mIsCrtTbl As Boolean: mIsCrtTbl = (J = 0)
'    If jj.Bld_OdbcTwoQry_BySqlDsn(pAySelSql(J), pAyDsn(J), mIsCrtTbl, pNmtTar, pFbTar, mNmq1, mNmq2, True) Then ss.A 4:Goto E
'Next
'Exit Function
'    Tbl_FmAySql_ByDsn = True
'End Function
#If Tst Then
Function Crt_Tbl_FmAySql_ByDsn_Tst() As Boolean
Const cSub$ = "Crt_Tbl_FmAySql_ByDsn_Tst"
Dim N%: N = 1
Dim mNmtTar$, mFbOupTbl$
ReDim mAyDsn$(0 To N)
ReDim mAySql$(0 To N)
Dim mCase As Byte: mCase = 1
Dim mResult As Boolean
Select Case mCase
Case 1
    mNmtTar$ = "tmpIIC"
    '
    mAyDsn(0) = "CHPROD_BPCSF"
    mAySql(0) = "Select 'CH' AS SRC, IIC.* from IIC"
    '
    mAyDsn(1) = "FEPROD_RBPCSF"
    mAySql(1) = "Select 'FE' AS SRC, IIC.* from IIC"
End Select
Debug.Print jj.Crt_Tbl_FmAySql_ByDsn(mNmtTar, mAyDsn, mAySql, mFbOupTbl)
End Function
#End If
Function Crt_TblVer(Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: create table tblVer in {pDb}
Const cSub$ = "Crt_TblVer"
jj.Set_Silent
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
If jj.Chk_Struct_Tbl("TblVer", "Ver", mDb) Then
    If jj.Dlt_Tbl("TblVer", mDb) Then ss.A 1: GoTo E
    If jj.Crt_Tbl_FmLoFld("tblVer", "Ver Date", , , mDb) Then ss.A 2: GoTo E
    With mDb.TableDefs("TblVer").OpenRecordset
        .AddNew
        !Ver = Now
        .Update
        .Close
    End With
Else
    If jj.Set_Ver Then ss.A 3: GoTo E
End If
GoTo X
R: ss.R
E: Crt_TblVer = True: ss.B cSub, cMod, "pDb", jj.ToStr_Db(pDb)
X:
    jj.Set_Silent_Rst
End Function
#If Tst Then
Function Crt_TblVer_Tst() As Boolean
If jj.Crt_TblVer Then Stop
End Function
#End If
Function Crt_TblTrc(Optional pDb As DAO.Database) As Boolean
'Aim: create table tblHst_Tp & tblHst_TpStpsin {pDb}
Const cSub$ = "Crt_TblTrc"
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
jj.Set_Silent
On Error GoTo R
If jj.Chk_Struct_Tbl("Trc", "Trc,Sess,ReqNo,NmLgs,Lp,Lv,DteBeg,DteEnd", mDb) Then
    If jj.Dlt_Tbl("Trc") Then ss.A 1: GoTo E     'Trc Step
    If jj.Crt_Tbl_FmLoFld("Trc", "Trc Long, Sess Long, TrcNo Long, NmLgs Text 50, Lp Text 255, Lv Memo,DteBeg Date,DteEnd Date", 1, , mDb) Then ss.A 2: GoTo E
End If
'TrcS=Trace Steps
If jj.Chk_Struct_Tbl("TrcS", "Trc,StepNo,NmLgc,Lp,Lv,DteBeg,DteEnd", mDb) Then
    If jj.Dlt_Tbl("TrcS") Then ss.A 1: GoTo E
    If jj.Crt_Tbl_FmLoFld("TrcS", "Trc Long, StepNo Int, NmLgc Text 50, Lp Text 255, Lv Memo,DteBeg Date,DteEnd Date", 2, , mDb) Then ss.A 3: GoTo E
End If
GoTo X
R: ss.R
E: Crt_TblTrc = True: ss.B cSub, cMod, "pDb", jj.ToStr_Db(pDb)
X:
    jj.Set_Silent
End Function
#If Tst Then
Function Crt_TblTrc_Tst() As Boolean
If jj.Crt_TblTrc Then Stop
End Function
#End If
Function Crt_TblPrm(Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: create table tblPrm in {pDb}
Const cSub$ = "Crt_TblPrm"
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
jj.Set_Silent
If jj.Chk_Struct_Tbl("tblPrm", "Trc,NmLgc,Lp,Lv", mDb) Then
    On Error GoTo R
    If jj.Dlt_Tbl("tblPrm", mDb) Then ss.A 1: GoTo E
    mDb.Execute ("Create table tblPrm (Trc Long, NmLgc Text(50), Lp Text(255), Lv Memo)")
    mDb.TableDefs.Refresh
End If
GoTo X
R: ss.R
E: Crt_TblPrm = True: ss.B cSub, cMod, "pDb", jj.ToStr_Db(pDb)
X:
    jj.Set_Silent_Rst
End Function
#If Tst Then
Function Crt_TblPrm_Tst() As Boolean
If jj.Crt_TblPrm Then Stop
End Function
#End If
Function Crt_Pk(pNmt$, pLnFld$, Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: Create PrimaryKey on {pNmt} by {pLnFld}
Const cSub$ = "Crt_Pk"
On Error GoTo R
If jj.Dlt_Idx(pNmt, "PrimaryKey", pDb) Then ss.A 1: GoTo E
Dim mSql$: mSql = jj.Fmt_Str("Create Index PrimaryKey on {0} ({1}) Primary", jj.Q_SqBkt(pNmt), pLnFld)
If jj.Run_Sql_ByDbExec(mSql, pDb) Then ss.A 2: GoTo E
Exit Function
R: ss.R
E: Crt_Pk = True: ss.B cSub, cMod, "pNmt,pLnFld,pDb", pNmt, pLnFld, jj.ToStr_Db(pDb)
End Function
Function Crt_Idx(pNmt$, pNmIdx$, pLnFld$, Optional pIsUniq As Boolean = False, Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: Create {pIdx} on {pNmt} by {pLnFld}
Const cSub$ = "Crt_Idx"
On Error GoTo R
If jj.Dlt_Idx(pNmt, pNmIdx, pDb) Then ss.A 1: GoTo E
Dim mSql$: mSql = jj.Fmt_Str("Create {0}Index {1} on {2} ({3})", IIf(pIsUniq, "UNIQUE ", ""), pNmIdx, jj.Q_SqBkt(pNmt), pLnFld)
If jj.Run_Sql_ByDbExec(mSql, pDb) Then ss.A 2: GoTo E
Exit Function
R: ss.R
E: Crt_Idx = True: ss.B cSub, cMod, "pNmt,pNmIdx,pLnFld,pIsUniq,pDb", pNmt, pNmIdx, pLnFld, pIsUniq, jj.ToStr_Db(pDb)
End Function
#If Tst Then
Function Crt_Idx_Tst() As Boolean
If jj.Crt_Tbl_FmLoFld("aa", "aa text 10, bb text 10") Then Stop
If jj.Crt_Idx("aa", "U01", "aa,bb") Then Stop
If jj.Crt_Idx("aa", "U01", "bb,aa") Then Stop
DoCmd.OpenTable "aa", acViewDesign
End Function
#End If
Function Crt_Tbl_FmTblF(Optional pNmtTblF$ = "#TblF") As Boolean
'Aim: Create all tables as defined in {pNmtTblF} to the Fb
'     #TblF: Pth,NmMdb,NPk,StopAutoInc,TblAtr,NmTbl,SnoTblF,NmFld,TypDao,FldLen,FmtTxt,IsReq,IsAlwZerLen,VdtTxt,VdtRul,DftVal
Const cSub$ = "Crt_Tbl_FmTblF"
On Error GoTo R
Dim mAyFld() As DAO.Field
If jj.Chk_Struct_Tbl(pNmtTblF, "Pth,NmMdb,NPk,StopAutoInc,TblAtr,NmTbl,SnoTblF,NmFld,TypDao,FldLen,FmtTxt,IsReq,IsAlwZerLen,VdtTxt,VdtRul,DftVal") Then ss.A 1: GoTo E
Dim mNmtTblF$: mNmtTblF = jj.Q_S(pNmtTblF, "[]")
Dim mAyFb$(): If jj.Fnd_AyVFmSql(mAyFb, "Select Distinct Pth & NmMdb from " & mNmtTblF) Then ss.A 1: GoTo E
Dim iFb%
For iFb = 0 To jj.Siz_Ay(mAyFb) - 1
    Dim mDb As DAO.Database: If jj.Opn_Db_RW(mDb, mAyFb(iFb)) Then ss.A 2: GoTo E
    Dim mAyNPk() As Byte, mAyStopAutoInc() As Boolean, mAyTblAtr&(), mAnt$()
    Dim mSql$: mSql = jj.Fmt_Str("Select Distinct NPk,StopAutoInc,TblAtr,NmTbl from {0} where Pth & NmMdb='{1}' order by NmTbl", mNmtTblF, mAyFb(iFb))
    If jj.Fnd_LoAyV_FmSql(mSql, "NPk,StopAutoInc,TblAtr,NmTbl", mAyNPk, mAyStopAutoInc, mAyTblAtr, mAnt) Then ss.A 2: GoTo E
    Dim iNmt%
    For iNmt = 0 To jj.Siz_Ay(mAnt) - 1
        jj.Shw_Sts "Creating Table " & mAnt(iNmt) & " ..."
        mSql = jj.Fmt_Str("Select NmFld,TypDao,FldLen,FmtTxt,IsReq,IsAlwZerLen,VdtTxt,VdtRul,DftVal from {0} where NmTbl='{1}' order by SnoTblF", mNmtTblF, mAnt(iNmt))
        Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, mSql) Then ss.A 3: GoTo E
        Dim J%: J = 0
        With mRs
            While Not .EOF
                ReDim Preserve mAyFld(J)
                If jj.Crt_Fld_FmRsTblF(mAyFld(J), mRs) Then ss.A 4: GoTo E
                J = J + 1
                .MoveNext
            Wend
            .Close
        End With
        If jj.Crt_Tbl_FmAyFld(mAnt(iNmt), mAyFld, mAyNPk(iNmt), mAyTblAtr(iNmt), mDb, mAyStopAutoInc(iNmt)) Then ss.A 5: GoTo E
        
        'Add ZerRec
        If mAyNPk(iNmt) = 1 Then
            If Left(mAnt(iNmt), 3) <> "$Ty" Or mAnt(iNmt) = "$TypDta" Then
                Dim mAnFld$(): If jj.Fnd_AnFld_ReqTxt(mAnFld, mAnt(iNmt), mDb) Then ss.A 3: GoTo E
                Dim mLnFld$, mLnVal$
                mLnFld = "": mLnVal = ""
                Dim I%, NFld%: NFld = jj.Siz_Ay(mAnFld)
                For I = 0 To NFld - 1
                    mLnFld = mLnFld & "," & mAnFld(I)
                    mLnVal = mLnVal & ",'-'"
                Next
                mSql = jj.Fmt_Str("Insert into [${0}] ({0}{1}) values (0{2})", mID(mAnt(iNmt), 2), mLnFld, mLnVal)
                If jj.Run_Sql_ByDbExec(mSql, mDb) Then ss.A 4: GoTo E
            End If
        End If
    Next
    mDb.Close
Next
GoTo X
R: ss.R
E: Crt_Tbl_FmTblF = True: ss.B cSub, cMod, "pNmtTblF", pNmtTblF
X:
    jj.Cls_Db mDb
    jj.Cls_Rs mRs
    jj.Clr_Sts
End Function
#If Tst Then
Function Crt_Tbl_FmTblF_Tst() As Boolean
If jj.Crt_Tbl_FmLnkNmt("p:\workingdir\pgmobj\JMtcDb.mdb", "#TblF") Then Stop: GoTo E
If jj.Crt_Tbl_FmTblF Then Stop: GoTo E
Exit Function
E: Crt_Tbl_FmTblF_Tst = True
End Function
#End If
Function Crt_Tbl_FmAmFld(pNmt$, pAmFld() As tMap, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Crt_Tbl_FmAmFld"
On Error GoTo R
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim mNmt$: mNmt = jj.Rmv_SqBkt(pNmt)
Dim mTbl As DAO.TableDef: Set mTbl = mDb.CreateTableDef(mNmt)
Dim J%
For J = 0 To jj.Siz_Am(pAmFld) - 1
    Dim mFld As DAO.Field
    Dim mTypDao As DAO.DataTypeEnum, mLen As Byte
    If jj.Cv_TypDAO_FmFldDcl(mTypDao, mLen, pAmFld(J).F2) Then ss.A 1: GoTo E
    If mTypDao = dbText Then
        Set mFld = mTbl.CreateField(pAmFld(J).F1, mTypDao, mLen)
    Else
        Set mFld = mTbl.CreateField(pAmFld(J).F1, mTypDao)
    End If
    mTbl.Fields.Append mFld
Next
If jj.Dlt_Tbl(mNmt, pDb) Then ss.A 1: GoTo E
With mDb.TableDefs
    .Append mTbl
    .Refresh
End With
Exit Function
R: ss.R
E: Crt_Tbl_FmAmFld = True: ss.B cSub, cMod, "pNmt,pAmFld,pDb", pNmt, jj.ToStr_Am(pAmFld), jj.ToStr_Db(pDb)
End Function
Function Crt_Tbl_FmAyFld(pNmt$, pAyFld() As DAO.Field _
    , Optional pNPk As Byte = 0 _
    , Optional pTblAtr As DAO.TableDefAttributeEnum = 0 _
    , Optional pDb As DAO.Database = Nothing _
    , Optional pStopAutoInc As Boolean = False) As Boolean
'Aim: Delete then Create {pNmt} in {pDb} by {pAyFld} with {pTblAtr}.
Const cSub$ = "Crt_Tbl_FmAyFld"
Dim N%: N = jj.Siz_Ay(pAyFld): If N = 0 Then ss.A 1, "Siz of pAyFld is zero": GoTo E
On Error GoTo R
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim mNmt$: mNmt = jj.Rmv_SqBkt(pNmt)
Dim mTbl As DAO.TableDef: Set mTbl = mDb.CreateTableDef(mNmt, pTblAtr)
Do
    Dim mIdx As DAO.Index
    If pNPk > 0 Then
        Set mIdx = mTbl.CreateIndex("PrimaryKey")
        mIdx.Unique = True
        mIdx.Primary = True
    End If
    
    Dim J%
    If pNPk = 1 And Not pStopAutoInc Then
        If pAyFld(J).Type = dbLong Then pAyFld(J).Attributes = DAO.FieldAttributeEnum.dbAutoIncrField
    End If
    For J = 0 To jj.Siz_Ay(pAyFld) - 1
        'Update mIdx if pNPk>J
        If pNPk > J Then
            With pAyFld(J)
                Dim mFld_K As DAO.Field
                Select Case .Type
                Case DAO.DataTypeEnum.dbText:       Set mFld_K = mTbl.CreateField(.Name, .Type, .Size)
                'Case DAO.DataTypeEnum.dbCurrency:   ss.A 3, "Currency field should be included in PK", "The Field", pAyFld(J).Name: GoTo E
                Case Else:                          Set mFld_K = mTbl.CreateField(.Name, .Type)
                End Select
            End With
            mIdx.Fields.Append mFld_K
        End If
        If jj.Add_Fld(mTbl, pAyFld(J)) Then ss.A 2: GoTo E
    Next
    If pNPk > 0 Then mTbl.Indexes.Append mIdx
Loop Until True

'Create table from mTbl
If jj.Dlt_Tbl(mNmt, pDb) Then ss.A 1: GoTo E
With mDb.TableDefs
    .Append mTbl
    .Refresh
End With
GoTo X
R: ss.R
E: Crt_Tbl_FmAyFld = True: ss.B cSub, cMod, "pNmt,pAyFld,pNPk,pTblAtr,pDb", pNmt, "..", pNPk, pTblAtr, jj.ToStr_Db(pDb)
X:
End Function
Function Crt_Tbl_FmLoFld(pNmt$, pLoFld$ _
        , Optional pNPk As Byte = 0 _
        , Optional pTblAtr As DAO.TableDefAttributeEnum = 0 _
        , Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: Delete then Create {pDb}!{pNmt} by {pLoFld} with {pTblAtr}.
'     Format of pLoFld is xxx Text 10,....
'     Note: xxx may be in xx^xx format.  ^ means for space
'       TEXT,CURRENCY,LONG,INT,BYTE,DATE,SINGLE,DOUBLE,MEMO,YESNO
Const cSub$ = "Crt_Tbl_FmLoFld"
'Do build mTbl & mIdx
On Error GoTo R
Dim mAyFldDcl$(): mAyFldDcl = Split(pLoFld, cComma)
Dim N%: N = jj.Siz_Ay(mAyFldDcl)
ReDim mAyFld(N - 1) As DAO.Field
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim mTbl As New DAO.TableDef: ' Set mTbl = mDb.CreateTableDef(jj.Rmv_SqBkt(pNmt), pTblAtr)
Dim J%
For J = 0 To N - 1
    Dim mNmFld$, mTyp As DAO.DataTypeEnum, mLen As Byte
    If jj.Brk_FldDcl(mNmFld, mTyp, mLen, mAyFldDcl(J)) Then ss.A 1: GoTo E
    
    If mTyp = dbText Then
        Set mAyFld(J) = mTbl.CreateField(mNmFld, mTyp, mLen)
    Else
        Set mAyFld(J) = mTbl.CreateField(mNmFld, mTyp)
    End If
    Select Case mTyp
    Case DAO.DataTypeEnum.dbText, _
        DAO.DataTypeEnum.dbMemo: mAyFld(J).AllowZeroLength = True
    End Select
    If Right(mAyFldDcl(J), 4) = "AUTO" Then mAyFld(J).Attributes = DAO.FieldAttributeEnum.dbAutoIncrField
Next
'Create table from mTbl
If jj.Crt_Tbl_FmAyFld(pNmt, mAyFld, pNPk, pTblAtr, pDb) Then ss.A 3: GoTo E
Exit Function
R: ss.R
E: Crt_Tbl_FmLoFld = True: ss.B cSub, cMod, "pNmt,pLoFld,pNPk,pTblAtr,pDb,J", pNmt, pLoFld, pNPk, pTblAtr, jj.ToStr_Db(pDb), J
End Function
#If Tst Then
Function Crt_Tbl_FmLoFld_Tst() As Boolean
'If jj.Run_Sql("Create table aXa (bb NUMERIC)") Then Stop
Dim mFmLoFld$, mNmt$
Dim mDb As DAO.Database: If jj.Crt_Db(mDb, "c:\tmp\aa.mdb", True) Then Stop
Dim mCase As Byte: mCase = 2
Select Case mCase
Case 1
    mNmt$ = "XX"
    mFmLoFld = "aa Long, bb Int, cc currency 4,TT TEXT 10"
Case 2
    mNmt = "MSysIMEXSpecs"
    mFmLoFld = "SpecName Text 64" & _
        ", SpecId Auto" & _
        ", DateDelim Text 2" & _
        ", DateFourDigitYear YesNo" & _
        ", DateLeadingZeros YesNo" & _
        ", DecimalPoint Text 2" & _
        ", DateOrder Int" & _
        ", FieldSeparator Text 2" & _
        ", FileType Int" & _
        ", SpecType Byte" & _
        ", StartRow Long" & _
        ", TextDelim Text 2" & _
        ", TimeDelim Text 2"
    If jj.Crt_Tbl_FmLoFld(mNmt, mFmLoFld, 1, 2, mDb) Then ss.A 1: GoTo E
End Select
jj.Cls_Db mDb
If jj.Opn_CurDb(jj.g.gAcs, "c:\tmp\aa.mdb") Then Stop
jj.g.gAcs.Visible = True
Stop
GoTo X
E:
X: jj.Cls_CurDb jj.g.gAcs
End Function
#End If
'Function Crt_Tbl_FmTblPrm(oNmLgc$) As Boolean
''Aim: Use first record in tblPrm to create a [Input Table] & Inp#<NmLgc>_Prm and Return the NmLgc$
'''    tblPrm: Trc,NmLgc,Ln,Lv
'''    [Input Table] name   = Inp#{NmLgc}
'''                  fields = Use the names in {Ln} as field name.  Type will be Text or Memo, depend on the length of each value in {Lv}
'Const cSub$ = "Crt_Tbl_FmTblPrm"
'Dim mTrc&, mLm$
'If jj.Fnd_Prm_FmTblPrm(mTrc, oNmLgc, mLm) Then ss.a 1: GoTo E
'Dim mNmt_Inp$: mNmt_Inp = "Inp#" & oNmLgc & "_Prm"
'Dim mAm() As tMap: mAm=jj.Get_Am_ByLm( mLm) Then ss.a 2: GoTo E
'Dim mSqlx$, mSql_Ln$, mSql_Lv$, J%, L%
'For J = 0 To UBound(mAnPrm)
'    Dim mNm$: mNm = "[" & mAnPrm(J) & "]"
'    L = Len(mAyV(J)): If L = 0 Then L = 1
'    If L > 255 Then
'        mSqlx = jj.Add_Str(mSqlx, mNm & " Memo")
'    Else
'        mSqlx = jj.Add_Str(mSqlx, mNm & " Text(" & L & ")")
'    End If
'    mSql_Ln = jj.Add_Str(mSql_Ln, mNm)
'    mSql_Lv = jj.Add_Str(mSql_Lv, jj.Q_S(mAyV(J)))
'Next
'If jj.Dlt_Tbl(mNmt_Inp) Then ss.a 1: GoTo E
'If jj.Run_Sql(jj.Fmt_Str("Create table {0} ({1})", mNmt_Inp, mSqlx)) Then ss.a 2: GoTo E
'If jj.Run_Sql(jj.Fmt_Str("Insert into {0} ({1}) values ({2})", mNmt_Inp, mSql_Ln, mSql_Lv)) Then ss.a 1: GoTo E
'Exit Function
'R: ss.R
'E: Crt_Tbl_FmTblPrm = True: ss.B cSub, cMod
'End Function
'#If Tst Then
'Function Crt_Tbl_FmTblPrm_Tst() As Boolean
'Dim mNmLgc$
'If jj.Crt_Tbl_FmTblPrm(mNmLgc) Then Stop
'Debug.Print mNmLgc
'End Function
'#End If
Function Crt_Clip_ByRsXX(pFfnTp$, pRno&, pRs As DAO.Recordset, oWb As Workbook) As Boolean
'Aim: Use {pRs} to {pRno} of {pFfnTp} & make copy to clip board and return to {oWb} (Because, close oWb will lose format
Const cSub$ = "Crt_Clip_ByRs"
Set oWb = g.gXls.Workbooks.Open(pFfnTp)
Dim mWs As Worksheet: Set mWs = oWb.Sheets(1)
mWs.Range("A" & pRno).CopyFromRecordset pRs
Dim mAdrLasCell$: mAdrLasCell = mWs.Cells.SpecialCells(xlCellTypeLastCell).Address
Dim mRge As Range: Set mRge = mWs.Range("A1:" & mAdrLasCell)
mRge.Copy
Exit Function
E: Crt_Clip_ByRsXX = True
End Function
#If Tst Then
Function Crt_Clip_ByRs_Tst() As Boolean
Dim mFfnFm$: mFfnFm = jj.Sffn_Tp("RmdInvDet")
If jj.Crt_Tbl_FmLnkNmt("tmpBldOneRmd_Det", jj.Sffn_SessTp("DD", "GenRmd", 1)) Then Stop
Dim mRs As DAO.Recordset: Set mRs = CurrentDb.TableDefs("tmpBldOneRmd_Det").OpenRecordset
Dim mWb As Workbook: If jj.Crt_Clip_ByRsXX(mFfnFm, 3, mRs, mWb) Then Stop
Stop
jj.Cls_Wb mWb
End Function
#End If
Function Crt_Dir(pDir$) As Boolean
Const cSub$ = "Crt_Dir"
If VBA.Dir(pDir, vbDirectory) <> "" Then Exit Function
On Error GoTo R
VBA.MkDir pDir
If VBA.Dir(pDir, vbDirectory) = "" Then ss.A 1, "After MkDir, the dir not exist": GoTo E
Exit Function
R: ss.R
E: Crt_Dir = True: ss.B cSub, cMod, "pDir", pDir
End Function
Function Crt_Fb(pFb$, Optional pOvrWrt As Boolean = False) As Boolean
Const cSub$ = "Crt_Fb"
Dim mDb As DAO.Database: If jj.Crt_Db(mDb, pFb, pOvrWrt) Then ss.A 1: GoTo E
On Error GoTo R
mDb.Close
Exit Function
R: ss.R
E: Crt_Fb = True: ss.B cSub, cMod, "pFb,pOvrWrt", pFb, pOvrWrt
End Function
Function Crt_Db(oDb As DAO.Database, pFb$, Optional pOvrWrt As Boolean = False) As Boolean
Const cSub$ = "Crt_Db"
On Error GoTo R
If jj.Ovr_Wrt(pFb, pOvrWrt) Then ss.A 1: GoTo E
Set oDb = jj.g.gDbEng.CreateDatabase(pFb, DAO.dbLangGeneral)
Exit Function
R: ss.R
E: Crt_Db = True: ss.B cSub, cMod, "pFb,pOvrWrt", pFb, pOvrWrt
End Function
#If Tst Then
Function Crt_Db_Tst() As Boolean
Dim mDb As DAO.Database
If jj.Crt_Db(mDb, "c:\tmp\aa.mdb", True) Then Stop
End Function
#End If
Function Crt_PDF_FmFfnPS(pFfnPS$, pFfnPDF$) As Boolean
Const cSub$ = "Crt_PDF_FmFfnPS"
If Not jj.IsFfn(pFfnPS) Then ss.A 1: GoTo E
If Right(pFfnPS, 3) <> ".PS" Then ss.A 1, "Must be PostScript file": GoTo E
If jj.Dlt_Fil(pFfnPDF) Then ss.A 3: GoTo E
Dim mCmd$: mCmd = jj.Fmt_Str("""C:\Program Files\PDFCreator\PDFCreator.exe"" /IF""{0}"" /OF""{1}"" /DELETEIF", pFfnPS, pFfnPDF)
Shell mCmd, vbMaximizedFocus
Exit Function
R: ss.R
E: Crt_PDF_FmFfnPS = True: ss.B cSub, cMod, "pFfnPS$, pFfnPDF$", pFfnPS$, pFfnPDF$
End Function
Function Crt_PDF_FmWrd(pFfnWrd$, Optional pFfnPDF$ = "", Optional pKeepWrd As Boolean = False) As Boolean
Const cSub$ = "Crt_PDF_FmDoc"
Dim mWrd As Word.Document: If jj.Opn_Wrd_R(mWrd, pFfnWrd) Then ss.A 1: GoTo E
Dim mFfnn$: mFfnn = jj.Cut_Ext(pFfnWrd)
Dim mFfnPDF$: mFfnPDF = Fct.NonBlank(pFfnPDF, mFfnn & ".pdf")
Dim mFfnPS$: mFfnPS = mFfnn & ".ps"
On Error GoTo R
If jj.Set_PdfPrt(True) Then ss.A 2: GoTo E
mWrd.PrintOut False, , , mFfnPS
If jj.Set_PdfPrt(False) Then ss.A 3: GoTo E
jj.Cls_Wrd mWrd, False
If Not pKeepWrd Then jj.Dlt_Fil pFfnWrd
Crt_PDF_FmWrd = jj.Crt_PDF_FmFfnPS(mFfnPS, mFfnPDF)
Exit Function
R: ss.R
E: Crt_PDF_FmWrd = True: ss.B cSub, cMod, "pFfnWrd,pFfnPdf", pFfnWrd, pFfnPDF
X:
    jj.Cls_Wrd mWrd, True
End Function
Function Crt_PDF_FmWrd_Tst() As Boolean
Const cSub$ = "Crt_PDF_FmWrd_Tst"
jj.Dlt_Fil "c:\RmdLvl1.Pdf": jj.Dlt_Fil "c:\RmdLvl1.doc": If jj.Cpy_Fil(jj.Sffn_Tp("ReminderLvl1(English)", , ".doc"), "c:\RmdLvl1.doc") Then Stop: GoTo E
jj.Dlt_Fil "c:\RmdLvl2.Pdf": jj.Dlt_Fil "c:\RmdLvl2.doc": If jj.Cpy_Fil(jj.Sffn_Tp("ReminderLvl2(English)", , ".doc"), "c:\RmdLvl2.doc") Then Stop: GoTo E
jj.Dlt_Fil "c:\RmdLvl3.Pdf": jj.Dlt_Fil "c:\RmdLvl3.doc": If jj.Cpy_Fil(jj.Sffn_Tp("ReminderLvl3(English)", , ".doc"), "c:\RmdLvl3.doc") Then Stop: GoTo E
If jj.Crt_PDF_FmWrd("c:\RmdLvl1.doc") Then GoTo E
If jj.Crt_PDF_FmWrd("c:\RmdLvl2.doc") Then GoTo E
If jj.Crt_PDF_FmWrd("c:\RmdLvl3.doc") Then GoTo E
If jj.Opn_PDF("c:\RmdLvl1.pdf") Then ss.A 1: GoTo E
If jj.Opn_PDF("c:\RmdLvl2.pdf") Then ss.A 2: GoTo E
If jj.Opn_PDF("c:\RmdLvl3.pdf") Then ss.A 3: GoTo E
Exit Function
R: ss.R
E: Crt_PDF_FmWrd_Tst = True: ss.B cSub, cMod
End Function
Function Crt_PDF_FmXls(pFx$, Optional pFfnPDF$ = "", Optional pKeepXls As Boolean = False) As Boolean
Const cSub$ = "Crt_PDF_FmXls"
Dim mWb As Workbook: If jj.Opn_Wb_R(mWb, pFx) Then ss.A 1: GoTo E
Dim mFfnn$: mFfnn = jj.Cut_Ext(pFx)
Dim mFfnPDF$: mFfnPDF = Fct.NonBlank(pFfnPDF, mFfnn & ".pdf")
Dim mFfnPS$: mFfnPS = mFfnn & ".ps"
On Error GoTo R
If jj.Set_PdfPrt(True) Then ss.A 2: GoTo E
mWb.PrintOut , , , , , True, , mFfnPS
If jj.Set_PdfPrt(False) Then ss.A 3: GoTo E
jj.Cls_Wb mWb, False
If Not pKeepXls Then jj.Dlt_Fil pFx
Crt_PDF_FmXls = jj.Crt_PDF_FmFfnPS(mFfnPS, mFfnPDF)
Exit Function
R: ss.R
E: Crt_PDF_FmXls = True: ss.B cSub, cMod, "pFx,pFfnPdf", pFx, pFfnPDF
End Function
Function Crt_PDF_FmXls_Tst() As Boolean
Crt_PDF_FmXls_Tst = jj.Crt_PDF_FmXls("M:\07 ARCollection\ARCollection\PgmDoc.xls")
End Function
Function Crt_Pt(pWb As Workbook, pWsnPt$, pRgeNam$, _
    pPfLst_Row$, pPfLst_Col$, pPfLst_Dta$, _
    Optional pPfTotLst_Row$ = "", Optional pPfTotLst_Col$ = "") As Boolean
Const cSub$ = "Crt_Pt"
'Aim: Create a new Ws of name {pWsnPt} having a Pt from a data source as defined in name {pRgeNam}
'     pPfLst is in format [<<PfCaption>>:]<<PfNam>>
Dim mPt As PivotTable: Set mPt = pWb.PivotCaches.Add(xlDatabase, pRgeNam).CreatePivotTable("", "Pt_" & pWsnPt)
mPt.PivotCache.MissingItemsLimit = xlMissingItemsNone
mPt.Parent.Name = pWsnPt
Dim Ay$(), J%, mNmPf$, mPfCaption$
With mPt
    'Set Pf col
    Ay = Split(pPfLst_Col, cComma)
    Dim pF As PivotField
    For J = UBound(Ay) To LBound(Ay) Step -1
        If jj.Brk_ColonAs_ToCaptionNm(mPfCaption, mNmPf, Ay(J)) Then ss.A 1: GoTo E
        Set pF = .PivotFields(mNmPf)
        With pF
            .Orientation = xlColumnField
            .Position = 1
            .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            If mPfCaption <> mNmPf Then .Caption = mPfCaption
        End With
    Next

    'Set Pf Row
    Ay = Split(pPfLst_Row, cComma)
    For J = UBound(Ay) To LBound(Ay) Step -1
        If jj.Brk_ColonAs_ToCaptionNm(mPfCaption, mNmPf, Ay(J)) Then ss.A 2: GoTo E
        Set pF = .PivotFields(mNmPf)
        With pF
            .Orientation = xlRowField
            .Position = 1
            .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            If mPfCaption <> mNmPf Then .Caption = mPfCaption
        End With
    Next

    'Set Pf Data
    Ay = Split(pPfLst_Dta, cComma)
    For J = UBound(Ay) To LBound(Ay) Step -1
        If jj.Brk_Str_Both(mNmPf, mPfCaption, Ay(J), ":") Then ss.A 1: GoTo E
        Set pF = .PivotFields(mNmPf)
        With pF
            .Orientation = xlDataField
            .Position = 1
            .Function = xlSum
            If mPfCaption <> "" Then .Caption = mPfCaption
        End With
    Next

    'Set Data Fields as col
    With mPt.DataPivotField
        .Orientation = xlColumnField
        .Position = 1
    End With
End With
jj.Srt_Pt mPt
Exit Function
R: ss.R
E: Crt_Pt = True: ss.B cSub, cMod, "pWb,pWsnPt,pRgeNam,pPfLst_Row$,pPfLst_Col$,pPfLst_Dta$,pPfTotLst_Row$,pPfTotLst_Col$", jj.ToStr_Wb(pWb), pWsnPt, pRgeNam, pPfLst_Row$, pPfLst_Col$, pPfLst_Dta$, pPfTotLst_Row$, pPfTotLst_Col$
End Function
Function Crt_Qry_FmTbl(pNmt$) As Boolean
'Aim: Create all queries as defined in {pNmt}: Fb,NmQry,Sql
Const cSub$ = "Crt_Qry_FmTbl"
If jj.Chk_Struct_Tbl(pNmt, "Fb,NmQry,Sql") Then ss.A 1: GoTo E
Dim mFbLas$, mRs As DAO.Recordset
If jj.Opn_Rs(mRs, "Select * from [" & jj.Rmv_SqBkt(pNmt) & "] order by Fb,NmQry") Then ss.A 2: GoTo E
With mRs
    While Not .EOF
        If mFbLas <> !Fb Then
            mFbLas = !Fb
            Dim mDb As DAO.Database: Cls_Db mDb: If jj.Opn_Db_RW(mDb, mFbLas) Then ss.A 2: GoTo E
        End If
        If jj.Crt_Qry(!NmQry, !Sql, mDb) Then ss.A 3: GoTo E
        .MoveNext
    Wend
End With
GoTo X
R: ss.R
E: Crt_Qry_FmTbl = True: ss.B cSub, cMod, "pNmt", pNmt
X:
    jj.Cls_Db mDb
    jj.Cls_Rs mRs
End Function
#If Tst Then
Function Crt_Qry_FmTbl_Tst() As Boolean
If jj.Crt_Tbl_FmLoFld("#FBQry", "Fb Text 255,NmQry Text 50,Sql Memo") Then GoTo E
If jj.Run_Sql("Insert into [#FBQry] values ('C:\Tmp\aa.mdb','qry1','select * from Tbl1')") Then GoTo E
If jj.Run_Sql("Insert into [#FBQry] values ('C:\Tmp\aa.mdb','qry2','select * from Tbl1')") Then GoTo E
If jj.Run_Sql("Insert into [#FBQry] values ('C:\Tmp\aa.mdb','qry3','select * from Tbl1')") Then GoTo E
If jj.Run_Sql("Insert into [#FBQry] values ('C:\Tmp\aa.mdb','qry4','select * from Tbl1')") Then GoTo E
If jj.Crt_Fb("c:\Tmp\aa.mdb", True) Then GoTo E
If Crt_Qry_FmTbl("#FBQry") Then GoTo E
jj.g.gAcs.OpenCurrentDatabase "c:\tmp\aa.mdb"
jj.g.gAcs.Visible = True
Exit Function
E: Crt_Qry_FmTbl_Tst = True
End Function
#End If
Function Crt_Qry(pNmq$, Optional pSql$ = "", Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Crt_Qry"
On Error GoTo R
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim mNmq$: mNmq = jj.Rmv_SqBkt(pNmq)
With mDb
    If jj.IsQry(mNmq, mDb) Then
        If .QueryDefs(mNmq).Type = DAO.QueryDefTypeEnum.dbQSQLPassThrough Then
            .QueryDefs.Delete (mNmq)
            .CreateQueryDef mNmq
        End If
    Else
        .CreateQueryDef mNmq
    End If
    Dim mQry As DAO.QueryDef: Set mQry = .QueryDefs(mNmq)
    If pSql <> "" Then mQry.Sql = pSql
    .QueryDefs.Refresh
End With
Exit Function
R: ss.R
E: Crt_Qry = True: ss.B cSub, cMod, "pNmq,pSql,pDb", pNmq, pSql, jj.ToStr_Db(pDb)
End Function
Function Crt_Qry_ByDSN(pNmq$, pSql$, pDsn$, pReturnsRecrods As Boolean, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Crt_Qry_ByDSN"
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
If jj.Crt_Qry(pNmq, "Select *;", mDb) Then ss.A 1: GoTo E
Dim mQry As DAO.QueryDef: Set mQry = mDb.QueryDefs(pNmq)
On Error GoTo R
mQry.Connect = "ODBC;DSN=" & pDsn & ";"
mQry.Sql = pSql
If jj.Set_QryPrp_Bool(mQry, "ReturnsRecords", pReturnsRecrods) Then ss.A 2: GoTo E
If jj.Set_QryPrp(mQry, "ODBCTimeout", SysCfg_OdbcTimeOut) Then ss.A 3: GoTo E
Exit Function
R: ss.R
E: Crt_Qry_ByDSN = True: ss.B cSub, cMod, "pNmq,pSql,pDSN", pNmq, pSql, pDsn
End Function
Function Crt_Qry_ByDSN_Tst() As Boolean
'mSql = "Select SUM(Case When ICLAS IN ('57','07') Then 1 Else 0 end) AA , SUM(Case When ICLAS IN ('14','64') Then 1  Else 0 end) BB from IIC"
'If jj.Crt_Qry_ByDSN("qry", mSql, "FEPROD_RBPCSF") Then Stop
jj.Shw_DbgWin
Debug.Print "----ReturnsRecords True------"
Crt_Qry_ByDSN_Tst = jj.Crt_Qry_ByDSN("xxxy", "Update YY SET ICDES='11' WHERE ICLAS='06'", "FETEST_QGPL", True)
Debug.Print jj.ToStr_Prps(CurrentDb.QueryDefs("XXX").Properties, vbCrLf)
Debug.Print "----ReturnsRecords False ------"
Crt_Qry_ByDSN_Tst = jj.Crt_Qry_ByDSN("xxx", "Update YY SET ICDES='11' WHERE ICLAS='06'", "FETEST_QGPL", False)
Debug.Print jj.ToStr_Prps(CurrentDb.QueryDefs("XXX").Properties, vbCrLf)
Debug.Print "----ReturnsRecords True------"
Crt_Qry_ByDSN_Tst = jj.Crt_Qry_ByDSN("xxx", "Update YY SET ICDES='11' WHERE ICLAS='06'", "FETEST_QGPL", True)
Debug.Print jj.ToStr_Prps(CurrentDb.QueryDefs("XXX").Properties, vbCrLf)
Debug.Print "----ReturnsRecords False ------"
Crt_Qry_ByDSN_Tst = jj.Crt_Qry_ByDSN("xxx", "Update YY SET ICDES='11' WHERE ICLAS='06'", "FETEST_QGPL", False)
Debug.Print jj.ToStr_Prps(CurrentDb.QueryDefs("XXX").Properties, vbCrLf)
End Function
Function Crt_Rel_FmTbl(pNmt$) As Boolean
'Aim: Create Relation for each record in {pNmt}: Fb,NmTbl,NmTblTo,RelNo,IsCascadeDlt,IsCascadeUpd,LmFld
Const cSub$ = "Crt_Rel_FmTbl"
If jj.Chk_Struct_Tbl(pNmt, "Fb,NmTbl,NmTblTo,RelNo,IsCascadeDlt,IsCascadeUpd,LmFld") Then ss.A 1: GoTo E
On Error GoTo R
Dim mNmt$: mNmt = jj.Q_SqBkt(pNmt)
Dim mAyFb$(): If jj.Fnd_AyVFmSql(mAyFb, "Select Distinct Fb from " & mNmt) Then ss.A 2: GoTo E
Dim J%
For J = 0 To jj.Siz_Ay(mAyFb) - 1
    Dim mDb As DAO.Database: If jj.Opn_Db_RW(mDb, mAyFb(J)) Then ss.A 3: GoTo E
    Dim mRs As DAO.Recordset, mSql$
    mSql = jj.Bld_SqlSel( _
        "NmTbl,NmTblTo,RelNo,IsCascadeDlt,IsCascadeUpd,LmFld" _
        , mNmt _
        , "Fb='" & mAyFb(J) & "'" _
        , "NmTbl,RelNo")
    If jj.Opn_Rs(mRs, mSql) Then ss.A 4: GoTo E
    With mRs
        While Not .EOF
            If jj.Crt_Rel(!NmTbl & "R" & Format(!RelNo, "00"), "$" & !NmTbl, "$" & !NmTblTo, !LmFld, True, !IsCascadeUpd, !IsCascadeDlt, mDb) Then ss.A 5: GoTo E
            .MoveNext
        Wend
        .Close
    End With
    jj.Cls_Db mDb
Next
GoTo X
R: ss.R
E: Crt_Rel_FmTbl = True: ss.B cSub, cMod, "pNmt", pNmt
X:
    jj.Cls_Rs mRs
    jj.Cls_Db mDb
End Function
Function Crt_Rel(pNmRel$, pNmtFm$, pNmtTo$, pLmFld$ _
    , Optional pIsIntegral As Boolean = False, Optional pIsCascadeUpd As Boolean = False, Optional pIsCascadeDlt As Boolean = False, Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: Create a relation. {pLmFld} is format of xx=yy,cc,dd=ee
Const cSub$ = "Crt_Rel"
On Error GoTo R
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
If jj.IsRel(pNmRel) Then ss.A 1: GoTo E
Dim mAm() As tMap: mAm = jj.Get_Am_ByLm(pLmFld)
If jj.Siz_Am(mAm) = 0 Then ss.A 3, "pLmFld given 0 siz Am()": GoTo E
On Error GoTo R
Dim mRelAtr As DAO.RelationAttributeEnum
If Not pIsIntegral Then mRelAtr = dbRelationDontEnforce
If pIsCascadeUpd Then mRelAtr = mRelAtr Or dbRelationUpdateCascade
If pIsCascadeDlt Then mRelAtr = mRelAtr Or dbRelationDeleteCascade
Dim mRel As DAO.Relation: Set mRel = mDb.CreateRelation(pNmRel, pNmtFm, pNmtTo, mRelAtr)
Dim J%
For J = 0 To jj.Siz_Am(mAm) - 1
    With mAm(J)
        mRel.Fields.Append mRel.CreateField(.F1)
        mRel.Fields(.F1).ForeignName = .F2
    End With
Next
mDb.Relations.Append mRel
Exit Function
R: ss.R
E: Crt_Rel = True: ss.B cSub, cMod, "pNmRel, pNmtFm, pNmtTo, pLmFld, pIsIntegral, pIsCascadeUpd, pIsCascadeDlt", pNmRel, pNmtFm, pNmtTo, pLmFld, pIsIntegral, pIsCascadeUpd, pIsCascadeDlt
End Function
Function Crt_Rel_Tst() As Boolean
Crt_Rel_Tst = jj.Crt_Rel("xxx#xx", "0Rec", "1Rec", "x", True, True, True)
End Function
Function Crt_Rge_Col(pWs As Worksheet, pCno1 As Byte, pCno2 As Byte) As Range
Set Crt_Rge_Col = pWs.Range(jj.Cv_Cno2Col(pCno1) & ":" & jj.Cv_Cno2Col(pCno2))
End Function
Function Crt_Rge_Fm1Pt2N(oRge As Range, pWs As Worksheet, pR&, Pc, pNRow&, pNCol As Byte) As Boolean
Const cSub$ = "Crt_Rge_Fm1Pt2N"
On Error GoTo R
With pWs
    Dim mRge As Range: Set mRge = .Cells(pR, Pc)
    Set oRge = .Range(mRge, mRge.Cells(pNRow, pNCol))
End With
Exit Function
R: ss.R
E: Crt_Rge_Fm1Pt2N = True: ss.B cSub, cMod, "pWs,pR,pC,pNRow,pNCol", jj.ToStr_Ws(pWs), pR, Pc, pNRow, pNCol
End Function
#If Tst Then
Function Crt_Rge_Fm1Pt2N_Tst() As Boolean
Dim mRge As Range:
Dim mWb As Workbook: If jj.Crt_Wb(mWb, "c:\tmp\aa.xls", True, "Sheet1") Then Stop: GoTo E
Dim mWs As Worksheet: Set mWs = mWb.Sheets(1)
If Crt_Rge_Fm1Pt2N(mRge, mWs, 4, "B", 2, 1) Then Stop: GoTo E
Debug.Print mRge.Address
GoTo X
E: Crt_Rge_Fm1Pt2N_Tst = True
X:
    jj.Cls_Wb mWb, , True
End Function
#End If
Function Crt_Rge_Fm2Pts(pWs As Worksheet, pR1&, pC1 As Byte, pR2&, pC2 As Byte) As Range
With pWs
    Set Crt_Rge_Fm2Pts = .Range(.Cells(pR1, pC1), .Cells(pR2, pC2))
End With
End Function
Function Crt_Rge_FmSq(pWs As Worksheet, pSq As tSq) As Range
With pWs
    Set Crt_Rge_FmSq = .Range(.Cells(pSq.r1, pSq.c1), .Cells(pSq.r2, pSq.c2))
End With
End Function
Function Crt_Rge_HLin(pWs As Worksheet, pRno&, pC1 As Byte, pC2 As Byte) As Range
Dim mC1$: mC1 = jj.Cv_Cno2Col(pC1)
Dim mC2$: mC2 = jj.Cv_Cno2Col(pC2)
Set Crt_Rge_HLin = pWs.Range(mC1 & pRno & ":" & mC2 & pRno)
End Function
Function Crt_Rge_NHCell(pWs As Worksheet, pAyCno() As Byte, pRno&) As Range
Dim mA$, J As Byte
mA = pWs.Cells(pRno, pAyCno(1)).Address
For J = 2 To UBound(pAyCno)
    mA = mA & cComma & pWs.Cells(pRno, pAyCno(J)).Address
Next
Set Crt_Rge_NHCell = pWs.Range(mA)
End Function
Function Crt_Rge_VLin(pWs As Worksheet, pCno As Byte, pR1&, pR2&) As Range
Dim mC$: mC = jj.Cv_Cno2Col(pCno)
Set Crt_Rge_VLin = pWs.Range(mC & pR1 & ":" & mC & pR2)
End Function
Function Crt_SessDta(pTrc&, Optional pFbTp$ = "") As Boolean
Const cSub$ = "Crt_SessDta"
'Aim: Create [Sess Sub Dir] under {pFbTp} & [mFbTp_Dta] if need.  If {pFbTp} is not given currentdb will be used.
Dim mFbTp$: If pFbTp = "" Then mFbTp = CurrentDb.Name Else mFbTp = pFbTp
Dim mDir$: mDir = Fct.Nam_DirNam(mFbTp) & Format(pTrc, "00000000") & "\"
If jj.Crt_Dir(mDir) Then ss.A 1: GoTo E
Dim mFbTp_Dta$: mFbTp_Dta = mDir & jj.Cut_Ext(Nam_FilNam(mFbTp)) & "_Dta.Mdb"
If VBA.Dir(mFbTp_Dta) = "" Then
    Dim mDb As DAO.Database
    If jj.Crt_Db(mDb, mFbTp_Dta) Then ss.A 2, "Cannot create the Tp's Data": GoTo E
    mDb.Close
End If
Exit Function
R: ss.R
E: Crt_SessDta = True: ss.B cSub, cMod, "pTrc,pFbTp", pTrc, pFbTp
End Function
#If Tst Then
Function Crt_SessDta_Tst() As Boolean
Const cFfn$ = "C:\aa.mdb"
If jj.Crt_SessDta(1) Then Stop
If jj.Crt_SessDta(1, cFfn) Then Stop
End Function
#End If
Function Crt_SubDtaSheet(pNmqFm$, pNmqTo$, pLnFldMst$, Optional pLnFldChd$ = "") As Boolean
Const cSub$ = "Crt_CrtRelForQry"
Dim mLnFldChd$: mLnFldChd = IIf(pLnFldChd = "", pLnFldMst, pLnFldChd)
Dim mQry As QueryDef: Set mQry = CurrentDb.QueryDefs(pNmqFm)
If jj.Set_QryPrp(mQry, "SubdatasheetName", pNmqTo) Then ss.A 1: GoTo E
If jj.Set_QryPrp(mQry, "LinkChildFields", mLnFldChd) Then ss.A 1: GoTo E
If jj.Set_QryPrp(mQry, "LinkMasterFields", pLnFldMst) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: Crt_SubDtaSheet = True: ss.B cSub, cMod, "pNmqFm$, pNmqTo$, pLnFldMst$, pLnFldChd$", pNmqFm$, pNmqTo$, pLnFldMst$, pLnFldChd$
End Function
Function Crt_SubDtaSheet_Tst() As Boolean
MsgBox jj.Crt_SubDtaSheet("qryARInq_1_LvlAsOf", "qryARInq_1_LvlCus", "InstId")
End Function
Function Crt_Tbl_FmDSN_Nmt(pDsn$, pNmt$, Optional pNmtTar$ = "" _
    , Optional pFbTar$ = "" _
    , Optional oDteBeg As Date, Optional oDteEnd As Date) As Boolean
If pNmtTar = "" Then pNmtTar = pNmt
Crt_Tbl_FmDSN_Nmt = jj.Crt_Tbl_FmDSN_Sql(pDsn, "Select * from " & pNmt, pNmtTar, pFbTar, oDteBeg, oDteEnd)
End Function
#If Tst Then
Function Crt_Tbl_FmDSN_Nmt_Tst() As Boolean
Dim mDteBeg As Date, mDteEnd As Date
If jj.Crt_Tbl_FmDSN_Nmt("FEPROD_RBPCSF", "iic", "tmpIIC", , mDteBeg, mDteEnd) Then Stop
Debug.Print mDteBeg
Debug.Print mDteEnd
End Function
#End If
Function Crt_Tbl_FmDSN_Sql(pDsn$, pSql$, pNmtTar$ _
    , Optional pFbTar$ = "" _
    , Optional oDteBeg As Date, Optional oDteEnd As Date) As Boolean
'Aim: Download Data to {pNmtTar} in {pFbTar} by {pSql} through {pDsn$}
Const cSub$ = "Crt_Tbl_FmDSN_Sql"
oDteBeg = Now
Dim Min$
If pFbTar <> "" Then
    If VBA.Dir(pFbTar) = "" Then If jj.Crt_Fb(pFbTar) Then ss.A 1: GoTo E
    Min = " IN '" & pFbTar & cQSng
End If
'If jj.Dlt_Tbl(pNmtTar, mDb) Then ss.A 1:Goto E
Dim mNmq$: mNmq = "qry" & Format(Now, "yyyymmddhhmmss")
If jj.Crt_Qry_ByDSN(mNmq, pSql, pDsn, True) Then ss.A 2: GoTo E
Dim mSql$: mSql = jj.Fmt_Str("Select * into {0} from {1}", jj.Q_S(pNmtTar, "[]") & Min, mNmq)
If jj.Run_Sql(mSql) Then ss.A 3: GoTo E
If jj.Dlt_Qry(mNmq) Then ss.A 4: GoTo E
oDteEnd = Now
Exit Function
R: ss.R
E: Crt_Tbl_FmDSN_Sql = True: ss.B cSub, cMod, "pDsn$, pSql$, pNmtTar$, pFbTar$, oDteBeg, oDteEnd", pDsn$, pSql$, pNmtTar$, pFbTar$, oDteBeg, oDteEnd
End Function
#If Tst Then
Function Crt_Tbl_FmDSN_Sql_Tst() As Boolean
Const cSub$ = "Crt_Tbl_FmDSN_Sql_Tst"
Dim mDsn$, mSql$, mNmtTar$, mFbTar$
Dim mRslt As Boolean, mCase As Byte
Dim mNRec&, mDteBeg As Date, mDteEnd As Date
jj.Shw_Dbg cSub, cMod
For mCase = 1 To 4
    Select Case mCase
    Case 1: mDsn = "FEPROD_RBPCSF": mSql = "Select * from IIC": mNmtTar = "IIC_Xls": mFbTar = "C:\aa.Mdb"
    Case 2: mDsn = "FEPROD_RBPCSF": mSql = "Select * from IIC": mNmtTar = "IIC_Txt": mFbTar = "C:\aa.Mdb"
    Case 3: mDsn = "FEPROD_RBPCSF": mSql = "Select * from IIC": mNmtTar = "IIC_Xls": mFbTar = ""
    Case 4: mDsn = "FEPROD_RBPCSF": mSql = "Select * from IIC": mNmtTar = "IIC_Txt": mFbTar = ""
    End Select
    mRslt = jj.Crt_Tbl_FmDSN_Sql(mDsn, mSql, mNmtTar, mFbTar, mDteBeg, mDteEnd)
    Debug.Print mCase; "-----------------------"
    Debug.Print jj.ToStr_LpAp(vbLf, "mRslt,mDsn,mSql,mNmtTar,mFbTar,mDteBeg,mDteEnd,mNRec", mRslt, mDsn, mSql, mNmtTar, mFbTar, mDteBeg, mDteEnd, mNRec)
Next
End Function
#End If
Function Crt_Tbl_FmDTF_Nmt(pIP$, pLib$, pNmt$, Optional pNmtTar$ = "", Optional pFbTar$ = "" _
    , Optional pIsByXls As Boolean = False _
    , Optional oDteBeg As Date, Optional oDteEnd As Date, Optional oNRec&) As Boolean
'Aim: Create {pNmtTar} in {pFbTar} from {pIP},{pLib},{pNmt} by meaning DTF download through {pIsByXls} or by Text
Crt_Tbl_FmDTF_Nmt = jj.Crt_Tbl_FmDTF_Sql(pIP, "Select * from " & pNmt, pNmtTar, pFbTar, pLib, pIsByXls, oDteBeg, oDteEnd, oNRec)
End Function
#If Tst Then
Function Crt_Tbl_FmDTF_Nmt_Tst() As Boolean
Const cSub$ = "Crt_Tbl_FmDTF_Nmt_Tst"
Dim mNRec&, mNmtTar$, mDteBeg As Date, mDteEnd As Date, mIsByXls As Boolean, mRslt
Dim mFbTar$
Dim mCase As Byte
jj.Shw_Dbg cSub, cMod
For mCase = 1 To 4
    Select Case mCase
        Case 1: mNmtTar = "IIC_ByXls": mIsByXls = True: mFbTar = "c:\aa.mdb"
        Case 2: mNmtTar = "IIC_ByTxt": mIsByXls = False: mFbTar = "c:\aa.mdb"
        Case 3: mNmtTar = "IIC_ByXls": mIsByXls = True: mFbTar = ""
        Case 4: mNmtTar = "IIC_ByTxt": mIsByXls = False: mFbTar = ""
    End Select
    mRslt = Crt_Tbl_FmDTF_Nmt("192.168.103.14", "RBPCSF", "IIC", mNmtTar, mFbTar, mIsByXls, mDteBeg, mDteEnd, mNRec)
    Debug.Print jj.ToStr_LpAp(vbTab, "mRslt, mFbTar, mIsByXls, mDteBeg, mDteEnd, mNRec", mRslt, mFbTar, mIsByXls, mDteBeg, mDteEnd, mNRec)
Next
End Function
#End If
Function Crt_Tbl_FmDTF_Sql(pIP$, pSql$, pNmtTar$, Optional pFbTar$ = "" _
    , Optional pLib$ = "RBPCSF" _
    , Optional pIsByXls As Boolean = False _
    , Optional oDteBeg As Date, Optional oDteEnd As Date, Optional oNRec& _
    ) As Boolean
'Aim: Create {pNmtTar} in {pFbTar} from {pIP},{pLib},{pSql} with time stamped & Rec count {oDteBeg,oDteEnd,oNRec&}.
Const cSub$ = "Crt_Tbl_FmDTF_Sql"
oDteBeg = Now

Dim mFfnDtf$: mFfnDtf = jj.Fnd_FfnDtf(pFbTar, pNmtTar)
If jj.Bld_Dtf(mFfnDtf, pSql, pIP, pLib, pIsByXls, True, oNRec) Then ss.A 1: GoTo E

If pIsByXls Then
    Dim mFx$: mFx = jj.Cut_Ext(mFfnDtf) & ".xls"
    If jj.Crt_Tbl_FmXls_n_FDF(mFx, pNmtTar, pFbTar) Then ss.A 2: GoTo E
Else
    Dim mFz$: mFz = jj.Cut_Ext(mFfnDtf) & ".txt"
    If jj.Crt_Tbl_FmTxt_n_FDF(mFz, pNmtTar, pFbTar) Then ss.A 3: GoTo E
End If
oDteEnd = Now
Exit Function
R: ss.R
E: Crt_Tbl_FmDTF_Sql = True: ss.B cSub, cMod, "pIP,pSql,pNmtTar,pFbTar,pLib,pIsByXls,oDteBeg,oDteEnd,oNRec", pIP, pSql, pNmtTar, pFbTar, pLib, pIsByXls, oDteBeg, oDteEnd, oNRec
End Function
#If Tst Then
Function Crt_Tbl_FmDTF_Sql_Tst() As Boolean
Const cSub$ = "Crt_Tbl_FmDTF_Sql_Tst"
Dim mNRec&, mNmt$, mDteBeg As Date, mDteEnd As Date, mIsByXls As Boolean, mRslt
Dim mFbTar$
Dim mCase As Byte
jj.Shw_Dbg cSub, cMod
For mCase = 3 To 3
    Select Case mCase
        Case 1: mNmt = "IIC_ByXls": mIsByXls = True: mFbTar = "c:\aa.mdb"
        Case 2: mNmt = "IIC_ByTxt": mIsByXls = False: mFbTar = "c:\aa.mdb"
        Case 3: mNmt = "IIC_ByXls": mIsByXls = True: mFbTar = ""
        Case 4: mNmt = "IIC_ByTxt": mIsByXls = False: mFbTar = ""
    End Select
    mRslt = jj.Crt_Tbl_FmDTF_Sql("192.168.103.13", "Select * from IIC where ICLAS='07'", mNmt, mFbTar, "BPCSF", mIsByXls, mDteBeg, mDteEnd, mNRec)
    Debug.Print jj.ToStr_LpAp(vbLf, "mRslt, mFbTar, mIsByXls, mDteBeg, mDteEnd, mNRec", mRslt, mFbTar, mIsByXls, mDteBeg, mDteEnd, mNRec)
Next
End Function
#End If
Function Crt_Tbl_FmCsv(pFfnCsv$, Optional pNmtNew$ = "", Optional pAcs As Access.Application = Nothing) As Boolean
Const cSub$ = "Crt_Tbl_FmCsv"
On Error GoTo R
Dim mAcs As Access.Application: Set mAcs = jj.Cv_Acs(pAcs)
Dim mDb As DAO.Database: Set mDb = mAcs.CurrentDb
Dim mNmtNew$: If pNmtNew = "" Then mNmtNew = Fct.Nam_FilNam(pFfnCsv) Else mNmtNew = pNmtNew
jj.Dlt_Tbl mNmtNew, mDb
mAcs.DoCmd.TransferText acImportDelim, , mNmtNew, pFfnCsv, True
GoTo X
R: ss.R
E: Crt_Tbl_FmCsv = True: ss.B cSub, cMod, "pFfnCsv,pNmtNew", pFfnCsv, pNmtNew
X: Set mDb = Nothing
End Function
#If Tst Then
Function Crt_Tbl_FmCsv_Tst() As Boolean
Dim mFfnCsv$
Dim mCase As Byte: mCase = 1
Select Case mCase
Case 1: mFfnCsv = "c:\Tmp\CsvChgTbl_20080518_175348(4).csv"
End Select
If Crt_Tbl_FmCsv(mFfnCsv, ">ChgTbl") Then Stop
DoCmd.OpenTable ">ChgTbl"
End Function
#End If
Function Crt_Tbl_FmLnkCsv(pFfnCsv$, Optional pNmtNew$ = "", Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Crt_Tbl_FmLnkCsv"
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim mNmtNew$: If pNmtNew = "" Then mNmtNew = Fct.Nam_FilNam(pFfnCsv) Else mNmtNew = pNmtNew
jj.Dlt_Tbl mNmtNew, mDb
Dim mTbl As New DAO.TableDef
On Error GoTo R
With mTbl
    Dim mDir$, mFnn$, mExt$
    Call jj.Brk_Ffn_To3Seg(mDir, mFnn, mExt, pFfnCsv)
    .Connect = jj.Fmt_Str("Text;DSN=Import Link Specification;FMT=Delimited;HDR=NO;IMEX=2;CharacterSet=936;DATABASE={0};TABLE={1}#{2}", mDir, mFnn, mID(mExt, 2))
    .Name = mNmtNew
    .SourceTableName = mFnn & mExt
    mDb.TableDefs.Append mTbl
End With
On Error GoTo 0
Exit Function
R: ss.R
E: Crt_Tbl_FmLnkCsv = True: ss.B cSub, cMod, "pFfnCsv,pNmtNew", pFfnCsv, pNmtNew
'Text;DSN=Import Link Specification;FMT=Delimited;HDR=NO;IMEX=2;CharacterSet=936;DATABASE=R:\Sales Simulation\Simulation\Import\2007_07_19 @01 55;TABLE=Import#Csv
    
End Function
Function Crt_Tbl_FmLnkCsv_Tst() As Boolean
Dim cFfnCsv$, cNmtNew$
'cFfnCsv$ = "R:\Sales Simulation\Simulation\Import\2007_07_19 @01 55\Import.Csv"
'cNmtNew$ = "tmpImp_Import"
'CrtTbl_FmLnkCsv_Tst = CrtTbl_FmLnkCsv(cFfnCsv, cNmtNew)
cFfnCsv$ = "R:\Sales Simulation\Simulation\Import\2007_07_19 @01 55\DataTotalEuro S01 BrandGp03-Nam\Val.csv"
cNmtNew$ = "tmpImp_Val"
Crt_Tbl_FmLnkCsv_Tst = jj.Crt_Tbl_FmLnkCsv(cFfnCsv, cNmtNew)
End Function
Function Crt_Tbl_FmLnkLdb(pFbLdb$, pLoadInstId&, pNmDb$, pLnt$) As Boolean
Const cSub$ = "Crt_Tbl_FmLnkLdb"
'Aim: Create a list of table in {pLnt} by referring {pFbLdb} & {pLoadInstId}
Dim mDb As DAO.Database: If jj.Opn_Db_R(mDb, pFbLdb) Then ss.A 1: GoTo E
Dim mLn_wQuote$: If jj.Q_Ln(mLn_wQuote, pLnt) Then ss.A 2: GoTo E
Dim mSql$: mSql = "Select" & _
" [SdirHom] & 'Mdb' & Format([MdbSno],'000') & '.Mdb' AS xFbTar," & _
" [NmHost] & '_' & [NmDb] & '_' & [Nmt]                 AS xNmt" & _
" from tblLdbHdr h inner join tblLdbDet d on h.LoadInstId=d.LoadInstId where h.LoadInstId=" & pLoadInstId & " and Nmt in (" & mLn_wQuote & ")"
With mDb.OpenRecordset(mSql)
    While Not .EOF
        Dim mNmt$:     mNmt = !xNmt
        Dim mFbTar$:   mFbTar = !xFbTar
        
        If jj.Crt_Tbl_FmLnkNmt(mFbTar, mNmt) Then ss.A 3: GoTo E
        .MoveNext
    Wend
    .Close
End With
GoTo X
R: ss.R
E: Crt_Tbl_FmLnkLdb = True: ss.B cSub, cMod, "pFbLdb,pLoadInstId,pNmDb,pLnt"
X:
    jj.Cls_Db mDb
End Function
#If Tst Then
Function Crt_Tbl_FmLnkLdb_Tst() As Boolean
Debug.Print jj.Crt_Tbl_FmLnkLdb("M:\07 ARCollection\ARCollection\WorkingDir\PgmObj\modLdmdb", 7, "RBPCSF", "IIM,IIC")
End Function
#End If
Function Crt_Tbl_FmLnkSetNmt(pFbSrc$, pSetNmt$, Optional pPfxNmt$ = "", Optional pSfxNmt$ = "", Optional pInDb As DAO.Database = Nothing) As Boolean
'Aim: Create pPfxNmt$ + pSetTbl$ + pSfxNmt in {pInDb} by linking {pFbSrc}!{ppSetNmt}
Const cSub$ = "Crt_Tbl_FmLnkpSetNmt"
jj.Shw_Sts "Create tables by linking [" & pSetNmt & "] in [" & pFbSrc & "] ...."
Dim mAnt$(), N%
Do
    Dim mDbSrc As DAO.Database: If jj.Cv_Db_FmFb(mDbSrc, pFbSrc) Then ss.A 1: GoTo E
    If jj.Fnd_Ant_BySetNmt(mAnt, pSetNmt, mDbSrc) Then jj.Cls_Db mDbSrc: ss.A 2: GoTo E
    jj.Cls_Db mDbSrc
    N = jj.Siz_Ay(mAnt)
Loop Until True

Dim J%
For J = 0 To N - 1
    If jj.Crt_Tbl_FmLnkNmt(pFbSrc, mAnt(J), pPfxNmt & mAnt(J) & pSfxNmt, pInDb) Then ss.A 3: GoTo E
Next
GoTo X
R: ss.R
E: Crt_Tbl_FmLnkSetNmt = True: ss.B cSub, cMod, "pFbSrc,pSetNmt,pPfxNmt,pSfxNmt,pInDb", pFbSrc, pSetNmt, pPfxNmt, pSfxNmt, jj.ToStr_Db(pInDb)
X: jj.Clr_Sts
End Function
#If Tst Then
Function Crt_Tbl_FmLnkSetNmt_Tst() As Boolean
Const cSub$ = "Crt_Tbl_FmLnkLnt_Tst"
Dim mFbSrc$, mSetNmt$, mPfxNmt$
Dim mResult As Boolean, mCase As Byte: mCase = 1
Select Case mCase
Case 1
    mSetNmt = "tbl*"
    mFbSrc = "D:\SPLHalfWayHouse\MPSDetail\VerNew@2007_01_04\WorkingDir\PgmObj\MPS_RfhFc.Mdb"
    mPfxNmt = "$"
End Select
mResult = jj.Crt_Tbl_FmLnkLnt(mFbSrc, mSetNmt, mPfxNmt)
jj.Shw_Dbg cSub, cMod, , "Result,mLnt,mSetNmt,mPfxNmt", mResult, mFbSrc, mSetNmt, mPfxNmt
End Function
#End If
Function Crt_Tbl_FmLnkLnt(pFbSrc$, pLnt$, Optional pLntNew$ = "", Optional pInDb As DAO.Database = Nothing) As Boolean
'Aim: Create NonBlank({pLntNew},{pLnt}) in {pInDb} by linking {pFbSrc}!{pLnt}
Const cSub$ = "Crt_Tbl_FmLnkLnt"
On Error GoTo R
Dim mAnt$():      If jj.Brk_Ln2Ay(mAnt, pLnt) Then ss.A 1: GoTo E
Dim mAntNew$():   If jj.Brk_Ln2Ay(mAntNew, Fct.NonBlank(pLntNew, pLnt)) Then ss.A 2: GoTo E
Dim N%: N = jj.Siz_Ay(mAnt)
Dim J%
For J = 0 To N - 1
    If jj.Crt_Tbl_FmLnkNmt(pFbSrc, mAnt(J), mAntNew(J), pInDb) Then ss.A 3: GoTo E
Next
Exit Function
R: ss.R
E: Crt_Tbl_FmLnkLnt = True: ss.B cSub, cMod, "pFbSrc,pLnt,pLntNew,pInDb", pFbSrc, pLnt, pLntNew, jj.ToStr_Db(pInDb)
End Function
#If Tst Then
Function Crt_Tbl_FmLnkLnt_Tst() As Boolean
Const cSub$ = "Crt_Tbl_FmLnkLnt_Tst"
Dim mLnt$, mFbSrc$, mLntNew$
Dim mResult As Boolean, mCase As Byte: mCase = 1
Select Case mCase
Case 1
    mLnt = "tblOdbcSql,tblFc"
    mFbSrc = "D:\SPLHalfWayHouse\MPSDetail\VerNew@2007_01_04\WorkingDir\PgmObj\MPS_RfhFc.Mdb"
    mLntNew = ""
End Select
mResult = jj.Crt_Tbl_FmLnkLnt(mFbSrc, mLnt, mLntNew)
jj.Shw_Dbg cSub, cMod, , "Result,mLnt,mFbSrc,mLntNew", mResult, mLnt, mFbSrc, mLntNew
End Function
#End If
Function Crt_Tbl_FmLnkNmt(pFb$, pNmt$, Optional pNmtNew$ = "", Optional pInDb As DAO.Database = Nothing) As Boolean
'Aim: Create NonBlank({pNmtNew},{pNmt}) in {pInDb} by linking {pFb}!{pNmt}
Const cSub$ = "Crt_Tbl_FmLnkNmt"
Dim mNmt$: mNmt = NonBlank(pNmtNew, pNmt)
Dim mNmtSrc$: mNmtSrc = pNmt
Dim mCnn$: mCnn = ";DATABASE=" & pFb ';DATABASE={pFb};TABLE={pNmt}
Crt_Tbl_FmLnkNmt = jj.Crt_Tbl_FmLnk(mNmt, mNmtSrc, mCnn, pInDb)
Exit Function
R: ss.R
E: Crt_Tbl_FmLnkNmt = True: ss.B cSub, cMod, "pFb,pNmt,pNmtNew,pInDb", pFb, pNmt, pNmtNew, jj.ToStr_Db(pInDb)
End Function
#If Tst Then
Function Crt_Tbl_FmLnkNmt_Tst() As Boolean
Dim mFb$: mFb = "c:\tmp\aa.mdb"
Dim mNmt$: mNmt = "tmpLnk_AA"
Dim mNmtNew$: mNmtNew = "$AA"
Dim mDb As DAO.Database: If jj.Crt_Db(mDb, mFb, True) Then Stop
If jj.Crt_Tbl_FmLoFld(mNmt, "aa text 10, bb int", , , mDb) Then Stop
If jj.Crt_Tbl_FmLnkNmt(mFb, mNmt, mNmtNew) Then Stop
End Function
#End If
Function Crt_Tbl_FmLnkAs400Dsn(pNmt$, Optional pLib$ = "RBPCSF", Optional pAs400Dsn$ = "FEPROD_RBPCSF", Optional pNmtNew$ = "", Optional pInDb As DAO.Database) As Boolean
'Aim: Create NonBlank({pNmtNew},{pLib}_{pNmt}) in {pInDb} by linking {pNmt} through {pAs400Dsn}.  Dsn must use *SQL Naming Convertion, ie
Const cSub$ = "Crt_Tbl_FmLnkAs400Dsn"
Dim mNmt$: mNmt = NonBlank(pNmtNew, pLib & "_" & pNmt)
Dim mCnn$: mCnn = jj.Fmt_Str("ODBC;DSN={0};", pAs400Dsn)
Dim mNmtSrc$: mNmtSrc = pLib & "." & pNmt
Crt_Tbl_FmLnkAs400Dsn = jj.Crt_Tbl_FmLnk(mNmt, mNmtSrc, mCnn, pInDb)
Exit Function
R: ss.R
E: Crt_Tbl_FmLnkAs400Dsn = True: ss.B cSub, cMod, "pNmt,pLib,pAs400Dsn", pNmt, pLib, pAs400Dsn
    Debug.Print "<--- Cannot link"
End Function
#If Tst Then
Function Crt_Tbl_FmLnkAs400Dsn_Tst() As Boolean
If jj.Crt_Tbl_FmLnkAs400Dsn("IIC", , , "xx") Then Stop
End Function
#End If
Function Crt_Tbl_FmLnk(pNmt$, pNmtSrc$, pCnn$, Optional pInDb As DAO.Database = Nothing) As Boolean
'Aim: Create {pNmt} in {pInDb} by linking {pNmtSrc} using {pCnn}
Const cSub$ = "Crt_Tbl_FmLnk"
Dim mInDb As DAO.Database: Set mInDb = jj.Cv_Db(pInDb)
If jj.Dlt_Tbl(pNmt, mInDb) Then ss.A 1: GoTo E
Dim mTbl As New DAO.TableDef
On Error GoTo R
With mTbl
    .Connect = pCnn
    .Name = pNmt
    .SourceTableName = pNmtSrc
    mInDb.TableDefs.Append mTbl
End With
Exit Function
R: ss.R
E: Crt_Tbl_FmLnk = True: ss.B cSub, cMod, "pNmt,pNmtSrc,pCnn,pInDb", pNmt, pNmtSrc, pCnn, jj.ToStr_Db(pInDb)
End Function
#If Tst Then
Function Crt_Tbl_FmLnk_Tst() As Boolean
Dim mNmt$:      mNmt = "A1"
Dim mNmtSrc$:   mNmtSrc = "a1.txt"
Dim mCnn$:      mCnn = "Text;DSN=A1;FMT=Fixed;HDR=NO;IMEX=2;CharacterSet=20127;DATABASE=c:\;TABLE=a1#txt"
Dim mDb As DAO.Database: If jj.Crt_Db(mDb, "c:\aa.mdb", True) Then Stop
If jj.Crt_Tbl_FmLnk(mNmt, mNmtSrc, mCnn, mDb) Then Stop
End Function
#End If
Function Crt_Tbl_FmLnkSetWs(pFx$, pSetWs$, Optional pPfxNmt$ = "", Optional pInDb As DAO.Database = Nothing) As Boolean
'Aim: Create table using pPfx + ws name in {pInDb} by linking {pFx}!{pSetWs}.
Const cSub$ = "Crt_Tbl_FmLnkSetWs"
jj.Shw_Sts "Linking [" & pFx & "]![" & pSetWs & "]......"
Dim mWb As Workbook: If jj.Opn_Wb_R(mWb, pFx) Then ss.A 1: GoTo E
Dim mAnWs$(): If jj.Fnd_AnWs_BySetWs(mAnWs, mWb, pSetWs) Then ss.A 2: GoTo E
Dim mCnn$: mCnn = jj.CnnStr_Xls(pFx)
Dim J%
For J = 0 To jj.Siz_Ay(mAnWs) - 1
    Dim mNmtSrc$: mNmtSrc = mAnWs(J) & "$"
    Dim mNmt$: mNmt = pPfxNmt & mAnWs(J)
    If jj.Crt_Tbl_FmLnk(mNmt, mNmtSrc, mCnn, pInDb) Then ss.A 3: GoTo E
Next
GoTo X
R: ss.R
E: Crt_Tbl_FmLnkSetWs = True: ss.B cSub, cMod, "pFx,pSetWs,pPfxNmt,pInDb", pFx, pSetWs, pPfxNmt, jj.ToStr_Db(pInDb)
X:
    jj.Clr_Sts
    jj.Cls_Wb mWb, False, True
End Function
Function Crt_Tbl_FmLnkWs(pFx$, Optional pNmWs$ = "", Optional pNmtNew$ = "", Optional pInDb As DAO.Database = Nothing) As Boolean
'Aim: Create NonBlank({pNmtNew},{pNmWs}) in {pInDb} by linking {pFx}!{pNmWs}.  If {pNmWs} is not given, use FileName(pFx).
Const cSub$ = "Crt_Tbl_FmLnkWs"
If pNmWs = "" Then pNmWs = jj.Cut_Ext(Fct.Nam_FilNam(pFx))
Dim mNmt$: mNmt = NonBlank(pNmtNew, pNmWs)
Dim mCnn$: mCnn = jj.CnnStr_Xls(pFx)
Dim mNmtSrc$: mNmtSrc = pNmWs & "$"
Crt_Tbl_FmLnkWs = jj.Crt_Tbl_FmLnk(mNmt, mNmtSrc, mCnn, pInDb)
Exit Function
R: ss.R
E: Crt_Tbl_FmLnkWs = True: ss.B cSub, cMod, "pFx,pNmWs,pNmtNew,pInDb", pFx, pNmWs, pNmtNew, jj.ToStr_Db(pInDb)
'Excel 8.0;HDR=YES;IMEX=2;DATABASE=D:\Data\MyDoc\Development\ISS\Imports\PO\PUR904 (On-Line).xls;TABLE='PUR904 (On-Line)'
End Function
#If Tst Then
Function Crt_Tbl_FmLnkWs_Tst() As Boolean
Const cSub$ = "Crt_Tbl_FmLnkWs_Tst"
Const cFx$ = "c:\tmp\aa.xls"
Const cNmWs$ = "aa"
Dim mWb As Workbook: If jj.Crt_Wb(mWb, cFx, True, cNmWs) Then Stop
Dim mWs As Worksheet: Set mWs = mWb.Sheets(1)
If jj.Set_Ws_ByLpAp(mWs, 1, "abc,def,xyz", 1, "a123", Now) Then Stop
If jj.Cls_Wb(mWb, True) Then Stop
If jj.Crt_Tbl_FmLnkWs(cFx, cNmWs) Then Stop
End Function
#End If
Function Crt_Tbl_FmLnkXls(pFx$, Optional pPfx$ = "", Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: Link all worksheets in {pFx} as tables in {pDb}
Const cSub$ = "Crt_Tbl_FmLnkXls"
jj.Shw_Sts "Create tables by linking [" & pFx & "]...."
Dim AnWs$():  If jj.Fnd_AnWs(AnWs, pFx) Then ss.A 1: GoTo E
Dim iNmWs, mA$
For Each iNmWs In AnWs
    Dim mNmWs$: mNmWs = iNmWs
    If jj.Crt_Tbl_FmLnkWs(pFx, mNmWs, pPfx & mNmWs, pDb) Then mA = jj.Add_Str(mA, mNmWs)
Next
If Len(mA) <> 0 Then ss.A 1, "Some ws {mA} in xls file cannot be linked", "mA", mA: GoTo E
GoTo X
R: ss.R
E: Crt_Tbl_FmLnkXls = True: ss.B cSub, cMod, "pFx,pPfx,pDb", pFx, pPfx, jj.ToStr_Db(pDb)
X:
    jj.Clr_Sts
End Function
Function Crt_Tbl_FmLnkXls_Tst() As Boolean
MsgBox jj.Crt_Tbl_FmLnkXls("c:\temp\LT\LT.xls")
End Function
Function Crt_Tbl_FmMgeNRec_To1Fld(pNmt$, Optional pSepChr$ = cComma, Optional pFillDta = False) As Boolean
'Aim: Create a table of name {pNmt}_Lst of 2 fields from the first 2 fields of {pNmt}.
'     The fields name of {pNmt}_Lst is same as the first 2 fields of {pNmt} with prefix [_Lst] in 2nd field
'     The 2nd field of {pNmt}_Lst is always memo no matter what field type of 2nd field of {pNmt}
'     The 1st field of {pNmt}_Lst will the PrimaryKey and this PrimaryKey will be created.
'     Create empty {pNmt}_Lst if pFillDta is false
Const cSub$ = "Crt_Tbl_FmMgeNRec_To1Fld"
Dim mF1$: mF1 = CurrentDb.TableDefs(pNmt).Fields(0).Name
Dim mF2$: mF2 = CurrentDb.TableDefs(pNmt).Fields(1).Name
Dim mSql$
mSql = jj.Fmt_Str("Select {0} into {1}_Lst from {1} where false", mF1, pNmt)
If jj.Run_Sql(mSql) Then ss.A 1: GoTo E
mSql = jj.Fmt_Str("Alter table {0}_Lst Add COLUMN {1}_Lst Memo", pNmt, mF2)
If jj.Run_Sql(mSql) Then ss.A 2: GoTo E
If Not pFillDta Then Exit Function
Dim mLasF1$, mF2Lst$
With CurrentDb.OpenRecordset(jj.Fmt_Str("Select {0},{1} from {2} order by {0},{1}", mF1, mF2, pNmt))
    If .AbsolutePosition <> -1 Then mLasF1 = .Fields(0).Value
    While Not .EOF
        If mLasF1 = .Fields(0).Value Then
            mF2Lst = jj.Add_Str(mF2Lst, CStr(Nz(.Fields(1).Value, "")), pSepChr)
        Else
            mSql = jj.Fmt_Str("Insert into {0}_Lst ({1},{2}_Lst) values ('{3}','{4}')", pNmt, mF1, mF2, mLasF1, mF2Lst)
            If jj.Run_Sql(mSql) Then ss.A 3: GoTo E
            mLasF1 = .Fields(0).Value
            mF2Lst = .Fields(1).Value
        End If
        .MoveNext
    Wend
    mSql = jj.Fmt_Str("Insert into {0}_Lst ({1},{2}_Lst) values ('{3}','{4}')", pNmt, mF1, mF2, mLasF1, mF2Lst)
    If jj.Run_Sql(mSql) Then ss.A 3: GoTo E
    .Close
End With
Exit Function
R: ss.R
E: Crt_Tbl_FmMgeNRec_To1Fld = True: ss.B cSub, cMod, "pNmt,pSepChr,pFillDta", pNmt, pSepChr, pFillDta
End Function
Function Crt_Tbl_FmMgeNRec_To1Fld_Tst() As Boolean
'tmpMPS_SKUFacParam is from MPSDetail.Mdb
Const cNmt$ = "tmpMPS_SKUFacParam"
Const cNmt_x$ = "tmpMPS_SKUFacParam_x"
Const cSub$ = "Crt_Tbl_FmMgeNRec_To1Fld_Tst"
DoCmd.CopyObject , cNmt_x, acTable, cNmt
Dim mSql$
mSql = jj.Fmt_Str("Update {0} set SKU_FacParam=Fac & ': ' & SKU_FacParam", cNmt_x)
If jj.Run_Sql(mSql) Then ss.A 1: GoTo E
mSql = jj.Fmt_Str("Alter table {0} Drop Column Fac", cNmt_x)
If jj.Run_Sql(mSql) Then ss.A 1: GoTo E

Dim mRslt As Boolean: mRslt = jj.Crt_Tbl_FmMgeNRec_To1Fld(cNmt_x, vbCrLf)
DoCmd.OpenTable cNmt_x & "_Lst"
Exit Function
R: ss.R
E: Crt_Tbl_FmMgeNRec_To1Fld_Tst = True: ss.B cSub, cMod
End Function
Function Crt_Tbl_FmTxt_n_TDF(pFz$, Optional pNmtTar$ = "", Optional pFbTar$ = "") As Boolean
Const cSub$ = "Crt_Tbl_FmTxt_n_TDF" ' TDF stands Table Decription File
'Aim: Create a table {pNmtTar} in {pFbTar} by import a text file {pFz} in {pDir} with schema.ini in same directory

End Function
Function Crt_Tbl_FmTxt_n_FDF(pFz$, Optional pNmtTar$ = "", Optional pFbTar$ = "") As Boolean
Const cSub$ = "Crt_Tbl_FmTxt_n_FDF"
'Aim: Create a table {pNmtTar} in {pFbTar} by import a text file {pFz} in {pDir} with schema.ini in same directory
Dim mFfnFdf$: mFfnFdf = jj.Cut_Ext(pFz) & ".Fdf"

'#1 Check if both {pFz} & [Fdf] exist
If VBA.Dir(mFfnFdf) = "" Then ss.A 1, "Fdf file not exist": GoTo E
If VBA.Dir(pFz) = "" Then
    Dim mF As Byte: If jj.Opn_Fil_ForOutput(mF, pFz) Then ss.A 2: GoTo E
    Close mF
End If

'#2 Build Schema.ini in {pDir}
If jj.Cv_Fdf2Schema(mFfnFdf) Then ss.A 3: GoTo E

'#3 Import
Dim mTarIn$, mCnnStr$, mSql$
If pFbTar <> "" Then
    If jj.Crt_Dir(pFbTar) Then If jj.Crt_Fb(pFbTar) Then ss.A 1: GoTo E
    mTarIn = " In '" & pFbTar & cQSng
End If

Dim mDir$, mFnn$, mExt$: If jj.Brk_Ffn_To3Seg(mDir, mFnn, mExt, pFz) Then ss.A 1: GoTo E

Dim mLnFld_Tar$: If jj.Cv_LnFld_FmFdf(mLnFld_Tar, mFfnFdf) Then ss.A 5: GoTo E
mCnnStr = "Text;Database=" & mDir
If pNmtTar = "" Then pNmtTar = jj.Cut_Ext(Fct.Nam_FilNam(pFz))
mSql = jj.Fmt_Str("Select {0} into [{1}]{2} from [{3}] in '' [{4}]", mLnFld_Tar, pNmtTar, mTarIn, mFnn & "#Txt", mCnnStr)

If jj.Run_Sql(mSql) Then ss.A 4: GoTo E

'#4 Dlt Txt, Fdf & Schema.ini if success
jj.Dlt_Fil pFz
jj.Dlt_Fil mFfnFdf
jj.Dlt_Fil mDir & "Schema.ini"
Exit Function
R: ss.R
E: Crt_Tbl_FmTxt_n_FDF = True: ss.B cSub, cMod, "pFz,pNmtTar,pFbTar", pFz, pNmtTar, pFbTar
End Function
#If Tst Then
Function Crt_Tbl_FmTxt_n_FDF_Tst() As Boolean
Const cFfnDtf$ = "C:\Tmp\IIC.dtf"
Const cNm$ = "IIC"
If jj.Bld_Dtf(cFfnDtf, "Select * from IIC", "192.168.103.14", , , True) Then Stop
If jj.Crt_Tbl_FmTxt_n_FDF(jj.Cut_Ext(cFfnDtf) & ".txt") Then Stop
End Function
#End If
Function Crt_Tbl_FmXls_n_FDF(pFx$, Optional pNmtTar$ = "", Optional pFbTar$ = "") As Boolean
 Const cSub$ = "Crt_Tbl_FmXls_n_FDF"
'Aim: Create a table {pNmtTar} in {pFbTar} by import an Xls file {pFx} with referring CutExt{pFx}.Fdf
Dim mFfnFdf$: mFfnFdf = jj.Cut_Ext(pFx) & ".Fdf"
If VBA.Dir(mFfnFdf) = "" Then ss.A 1, "Fdf file not exist (which is expected in same dir of pFx)": GoTo E

'if pFx not exist, build one from FDF
If VBA.Dir(pFx) = "" Then
    'Create pFx from Fdf
    If jj.Crt_Xls_FmFDF(pFx, mFfnFdf) Then ss.A 2: GoTo E
End If

'Set mIn_FbTar by pFbTar
Dim mIn_FbTar$
If pFbTar <> "" Then
    If VBA.Dir(pFbTar) Then If jj.Crt_Fb(pFbTar) Then ss.A 3: GoTo E
    mIn_FbTar = " In '" & pFbTar & cQSng
End If

'Import
Dim mLnFld_Tar$: If jj.Cv_LnFld_FmFdf(mLnFld_Tar, mFfnFdf) Then ss.A 4: GoTo E
Dim mCnnStr$: mCnnStr = jj.CnnStr_Xls(pFx)
Dim mNm$: mNm$ = jj.Cut_Ext(Fct.Nam_FilNam(pFx))
If pNmtTar = "" Then pNmtTar = mNm
Dim mSql$: mSql = jj.Fmt_Str("Select {0} into {1} from [{2}] in '' [{3}]", mLnFld_Tar, pNmtTar & mIn_FbTar, mNm & "$", mCnnStr)
If jj.Run_Sql(mSql) Then ss.A 4: GoTo E
jj.Dlt_Fil pFx
jj.Dlt_Fil mFfnFdf
Exit Function
R: ss.R
E: Crt_Tbl_FmXls_n_FDF = True: ss.B cSub, cMod, "pFx,pNmtTar,pFbTar", pFx, pNmtTar, pFbTar
End Function
#If Tst Then
Function Crt_Tbl_FmXls_n_FDF_Tst() As Boolean
Const cFfnDtf$ = "C:\Tmp\IIC.dtf"
Const cNm$ = "IIC"
If jj.Bld_Dtf(cFfnDtf, "Select * from IIC where ICLAS='xx'", "192.168.103.14", , , True, True) Then Stop
If jj.Crt_Tbl_FmXls_n_FDF(jj.Cut_Ext(cFfnDtf) & ".Xls") Then Stop
End Function
#End If
Function Crt_Tbl_ForEdtTbl(pNmtqSrc$, pNPk As Byte, Optional pNmtTar$ = "", Optional pStructOnly As Boolean = False) As Boolean
'Aim: Create table {mNmtTar} from {pNmtqSrc}.  {mNmtTar}'s content comes from {pNmtqSrc}.
'{mNmTar} fmt: first {pNPK} is same as {pNmtqSrc}, then a field [Change], then list of pair fields [xx] and [New xx]
Const cSub$ = "Crt_Tbl_ForEdtTbl"
If pNPk = 0 Then ss.A 1, "pNPK must > 0", , "pNmtqSrc,mNmTar", pNmtqSrc, pNmtTar: GoTo E
Dim mNmTar$: mNmTar = NonBlank(pNmtTar, "tmpEdt_" & pNmtqSrc)
If jj.Dlt_Tbl(mNmTar) Then ss.A 1: GoTo E

Dim mLnFld$
If jj.IsTbl(pNmtqSrc) Then
    mLnFld = jj.ToStr_Flds(CurrentDb.TableDefs(pNmtqSrc).Fields)
ElseIf jj.IsQry(pNmtqSrc) Then
    mLnFld = jj.ToStr_Flds(CurrentDb.QueryDefs(pNmtqSrc).Fields)
Else
    ss.A 1, "Given pNmtqSrc is not table or query": GoTo E
End If
Dim mAnFld$(): mAnFld = Split(mLnFld, cComma)
Dim A$: A = mAnFld(0)
Dim J%: For J = 1 To pNPk - 1
    A = ", " & mAnFld(J)
Next
A = A & ", " & "'' AS Changed"
Dim B$
For J = pNPk To UBound(mAnFld)
    A = A & ", [" & mAnFld(J) & "],'' as [New " & mAnFld(J) & "]"
    B = jj.Add_Str(B, "[New " & mAnFld(J) & "]=Null")
Next
A = A & ", '' As [Error During Import]"
Dim mSql$
mSql = jj.Fmt_Str("Select {0} into {1} from {2}", A, mNmTar, pNmtqSrc)
If pStructOnly Then
    If jj.Run_Sql(mSql & " Where False") Then ss.A 2: GoTo E
    Exit Function
End If
If jj.Run_Sql(mSql) Then ss.A 3: GoTo E
mSql = jj.Fmt_Str("Update {0} set {1}", mNmTar, B)
If jj.Run_Sql(mSql) Then ss.A 4: GoTo E
Exit Function
R: ss.R
E: Crt_Tbl_ForEdtTbl = True: ss.B cSub, cMod, "pNmtqSrc,mNmTar,pStructOnly", pNmtqSrc, mNmTar, pStructOnly
End Function
Function Crt_Tbl_ForEdtTbl_Tst() As Boolean
Const cSub$ = "Crt_Tbl_ForEdtTbl_Tst"
Dim mNmtqSrc$, mNmtTar$
Dim mRslt As Boolean, mCase As Byte: mCase = 2
Select Case mCase
Case 1
    mNmtqSrc = "tblUsr"
    mNmtTar = ""
Case 2
    mNmtqSrc = "tblCus"
    mNmtTar = ""
End Select
mRslt = jj.Crt_Tbl_ForEdtTbl(mNmtqSrc, 1, mNmtTar)
jj.Shw_Dbg cSub, cMod, , "mRslt, mNmtqSrc, mNmtTar", mRslt, mNmtqSrc, mNmtTar
End Function
Function Crt_Tbl_tmpXXX_Prm_By_qryOdbcXXX_0(pNmqsns$, Optional pLm$ = "") As Boolean
Const cSub$ = "Crt_"
Dim mNmtPrm$: mNmtPrm = "tmpOdbc" & pNmqsns & "_Prm"
Dim mAnq$(): If jj.Fnd_Anq_ByPfx(mAnq, "qryOdbc" & pNmqsns & "_0") Then ss.A 3: GoTo E
If jj.Run_Qry_ByAnq(mAnq, pLm) Then ss.A 4: GoTo E
If Not jj.IsTbl(mNmtPrm) Then ss.A 1, "Table mNmtPrm not exist", eRunTimErr, "mNmtPrm", mNmtPrm: GoTo E
Exit Function
R: ss.R
E: Crt_Tbl_tmpXXX_Prm_By_qryOdbcXXX_0 = True: ss.B cSub, cMod, "pNmqsns,pLm", pNmqsns, pLm
End Function
Function Crt_LgsTrc(oTrc&, pNmLgs$, Optional pLm$ = "") As Boolean
Const cSub$ = "Crt_LgsTrc"
Dim mSql$, mUsrID%: mUsrID = jj.UsrPrf_Usr
mSql = jj.Fmt_Str("Select Lp from tblLgs Where NmLgs='{0}'", pNmLgs)
Dim mLm$: If jj.Fnd_ValFmSql(mLm, mSql) Then ss.A 1, "Error in find record in tblTp by pNmLgs": GoTo E
If mLm <> pLm Then ss.A 2, "The Lp for pNmLgs in tblLgs is diff from given", "mLm from tblLgs", mLm: GoTo E
With CurrentDb.OpenRecordset("Select * from tblHst_Lgs")
    .AddNew
    oTrc = !Trc
    !NmLgs = pNmLgs
    !UsrId = mUsrID
    !Lm = pLm
    .Update
    .Close
End With
'Create records in tblHst_TpStps
mSql = jj.Fmt_Str( _
"INSERT INTO tblHst_LgsLgc ( Trc, NmLgc, Sno, Lp )" & _
" SELECT {0} AS Trc, tblLgsLgc.NmLgc, tblLgsLgc.Sno, tblLgsLgc.Lp" & _
" FROM tblLgs INNER JOIN tblLgsLgc ON tblLgs.Lgs = tblLgsLgc.Lgs" & _
" WHERE NmLgs='{1}'", _
oTrc, pNmLgs)
If jj.Run_Sql(mSql) Then ss.A 1: GoTo E
'Update Lm in each record in tblHst_TpStps
Dim mLnSubSet$
Stop
If pLm <> "" Then
    With CurrentDb.OpenRecordset("Select Lm from tblHst_LgsLgc where Trc=" & oTrc)
        While Not .EOF
            .Edit
            Dim mLmSubSet$: If jj.Cut_Lm(mLmSubSet, pLm, mLnSubSet) Then ss.A 1: GoTo E
            !Lm = mLmSubSet
            .Update
            .MoveNext
        Wend
        .Close
    End With
End If
Exit Function
R: ss.R
E: Crt_LgsTrc = True: ss.B cSub, cMod, "pNmLgs,pLm", pNmLgs, pLm
End Function
#If Tst Then
Function Crt_LgsTrc_Tst() As Boolean
Dim mTrc&: If jj.Crt_LgsTrc(mTrc, "MPS", "Env=FEPROD,Brand=TH") Then Stop
Debug.Print mTrc
End Function
#End If
Function Crt_TqRel(pNmtqFm$, pNmtqTo$, pLnkMstFlds$, Optional pLnkChdFlds$ = "") As Boolean
Const cSub$ = "Crt_CrtTqRel"
Dim mSubDsNm$
If jj.IsQry(pNmtqTo) Then
    mSubDsNm$ = "Query." & pNmtqTo
ElseIf jj.IsTbl(pNmtqTo) Then
    mSubDsNm$ = "Table." & pNmtqTo
Else
    ss.A 1, "Given pNmtqTo is not not table or query": GoTo E
End If

Dim mTypAcObj As Access.AcObjectType

If jj.IsTbl(pNmtqFm) Then
    mTypAcObj = acTable

ElseIf jj.IsQry(pNmtqFm) Then
    mTypAcObj = acQuery
Else
    ss.A 1, "Given pNmtqFm is not not table or query": GoTo E
End If

Dim mLnkChdFlds$: mLnkChdFlds = Fct.NonBlank(pLnkChdFlds, pLnkMstFlds)
If jj.Set_Prp(pNmtqFm, mTypAcObj, "SubdatasheetName", mSubDsNm$) Then ss.A 1: GoTo E
If jj.Set_Prp(pNmtqFm, mTypAcObj, "LinkMasterFields", pLnkMstFlds) Then ss.A 1: GoTo E
If jj.Set_Prp(pNmtqFm, mTypAcObj, "LinkChildFields", mLnkChdFlds) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: Crt_TqRel = True: ss.B cSub, cMod, "pNmtqFm,pNmtqTo,pLnkMstFlds,pLnkChdFlds", pNmtqFm, pNmtqTo, pLnkMstFlds, pLnkChdFlds
End Function
Function Crt_TqRel_Tst() As Boolean
Debug.Print jj.Crt_TqRel("qryCmp_04_0_Output", "qryCmp_05_1_FmLst_A_B", "K0_RPAN8;K1_RPDCT;K2_RPDOC")
End Function
Function Crt_Wb(oWb As Workbook, pFx$, Optional pOvrWrt As Boolean = False, Optional pNmWs$ = "ToBeDelete") As Boolean
Const cSub$ = "Crt_Wb"
If pFx = "" Then ss.A 1, "pFx cannot be empty": GoTo E
If jj.Ovr_Wrt(pFx, pOvrWrt) Then ss.A 1: GoTo E
Set oWb = gXls.Workbooks.Add()
If pNmWs <> "" Then
    While oWb.Sheets.Count > 1
        oWb.Sheets(1).Delete
    Wend
    oWb.Sheets(1).Name = pNmWs
End If
On Error GoTo R
oWb.SaveAs pFx
Exit Function
R: ss.R
E: Crt_Wb = True: ss.B cSub, cMod, "pFx,pOvrWrt,pNmWs", pFx, pOvrWrt, pNmWs
End Function
#If Tst Then
Function Crt_Wb_Tst() As Boolean
Dim mWb As Workbook
If jj.Crt_Wb(mWb, "c:\aa.xls", , "xxx") Then Stop
mWb.Application.Visible = True
End Function
#End If
Function Crt_Ws_FmWs(oWsTar As Worksheet, pWsFm As Worksheet, Optional pNmWsTo$ = "", Optional pWbTo As Workbook = Nothing) As Boolean
Const cSub$ = "Crt_Ws_FmWs"
'Aim: Copy {pWsFm} to a new {oWsTar}.
'Note: If {pWbTo} is given, the new Ws will be at end of {pWbTo}.  Otherwise, the oWsTar will be the same workbook as pWsFm.
'Note: If {pNmWsTo} is not given, the new Ws Name will use the {pWsFm}
If pNmWsTo = "" And jj.IsNothing(pWbTo) Then ss.A 1, "CpyWs must be given Either or both of {pNmWsTo}, {pWbTo}": GoTo E
'==Start
'Set {mWbTo} & {mNmWsTo}
Dim mWbTo As Workbook, mNmWsTo$
If jj.IsNothing(pWbTo) Then
    Set mWbTo = pWsFm.Parent
Else
    Set mWbTo = pWbTo
End If
If pNmWsTo = "" Then
    mNmWsTo = pWsFm.Name
Else
    mNmWsTo = pNmWsTo
End If
'Copy and Paste
pWsFm.Cells.Copy
Dim mNewWs As Worksheet: If jj.Add_Ws(mNewWs, mWbTo, mNmWsTo) Then ss.A 2: GoTo E
mNewWs.Cells.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
mNewWs.Cells.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Set oWsTar = mNewWs
Exit Function
R: ss.R
E: Crt_Ws_FmWs = True: ss.B cSub, cMod, "pWsFm,pNmWsTo,pWbTo", jj.ToStr_Ws(pWsFm), pNmWsTo, jj.ToStr_Wb(pWbTo)
End Function
Function Crt_Xls_FmFDF(pFxTar$, pFfnFdf$) As Boolean
Const cSub$ = "Crt_Xls_FmFDF"
'Aim: Create a 2 rows Xls {pFxTar} from {pFfnFDF}.  All numeric fields will set zero.
'PCFDF
'PCFT 16
'PCFO 1, 1, 5, 1, 1
'PCFL IID 20 4
'PCFL ICLAS 20 4
'PCFL ICDES 20 60
'PCFL ICGL 20 40
'PCFL ICCOGA 20 40
'PCFL ICTAX 20 10
'PCFL ICPPGL 20 40
'PCFL ICALGL 20 40
'PCFL ICALC1 20 20
'PCFL ICALC2 20 20
'PCFL ICALC3 20 20
'PCFL ICALC4 20 20
'PCFL ICALC5 20 20
'PCFL ICALG1 20 40
'PCFL ICALG2 20 40
'PCFL ICALG3 20 40
'PCFL ICALG4 20 40
'PCFL ICALG5 20 40
'PCFL ICALP1 2 7/2
'PCFL ICALP2 2 7/2
'PCFL ICALP3 2 7/2
'PCFL ICALP4 2 7/2
'PCFL ICALP5 2 7/2
'PCFL ICRETA 20 40
'PCFL ICCORA 20 40
'PCFL ICLMPC 2 7/2
'PCFL ICUMPC 2 7/2
On Error GoTo R
Dim mF As Byte: If jj.Opn_Fil_ForInput(mF, pFfnFdf) Then ss.A 1: GoTo E
Dim mL$, mA
Line Input #mF, mL: mA = "PCFDF":          If mL <> mA Then ss.A 2, "[" & mA & "] is expected", , "Current Line Value", mL: GoTo E
Line Input #mF, mL: mA = "PCFT 16":        If mL <> mA Then ss.A 3, "[" & mA & "] is expected", , "Current Line Value", mL: GoTo E
Line Input #mF, mL: mA = "PCFO 1,1,5,1,1": If mL <> mA Then ss.A 4, "[" & mA & "] is expected", , "Current Line Value", mL: GoTo E
Dim mWb As Workbook: If jj.Crt_Wb(mWb, pFxTar) Then ss.A 5, "Cannot create pFxTar": GoTo E
Dim mDir$, mFnn$, mExt$: If jj.Brk_Ffn_To3Seg(mDir, mFnn, mExt, pFxTar) Then ss.A 6: GoTo E
mWb.Sheets(1).Name = mFnn
Dim J%
For J = mWb.Worksheets.Count - 1 To 2 Step -1
    mWb.Worksheets(J).Delete
Next
J = 1
Dim mWs As Worksheet: Set mWs = mWb.Worksheets(1)
While Not EOF(mF)
    Line Input #mF, mL: If Left(mL, 5) <> "PCFL " Then ss.A 7, "[PCFL ] is expected", , "Current Line Value", mL: GoTo E
    Dim mX$(): mX = Split(mL)
    mWs.Cells(1, J).Value = mX(1)
    Select Case mX(2)
    Case "2": mWs.Cells(2, J).Value = 0
    End Select
    J = J + 1
Wend
jj.Cls_Wb mWb, True
Close #mF
Exit Function
R: ss.R
E: Crt_Xls_FmFDF = True: ss.B cSub, cMod, "pFxTar,pFfnFDF", pFxTar, pFfnFdf
End Function
Function Crt_Xls_FmFDF_Tst() As Boolean
Const cFfnFdf$ = "c:\aa.fdf"
Const cFx = "c:\aa.xls"
Dim mFno As Byte: If jj.Opn_Fil_ForOutput(mFno, cFfnFdf, True) Then Stop
Print #mFno, "PCFDF"
Print #mFno, "PCFT 16"
Print #mFno, "PCFO 1,1,5,1,1"
Print #mFno, "PCFL IID 20 4"
Print #mFno, "PCFL ICLAS 20 4"
Print #mFno, "PCFL ICDES 20 60"
Print #mFno, "PCFL ICGL 20 40"
Print #mFno, "PCFL ICCOGA 20 40"
Print #mFno, "PCFL ICTAX 20 10"
Print #mFno, "PCFL ICPPGL 20 40"
Print #mFno, "PCFL ICALGL 20 40"
Print #mFno, "PCFL ICALC1 20 20"
Print #mFno, "PCFL ICALC2 20 20"
Print #mFno, "PCFL ICALC3 20 20"
Print #mFno, "PCFL ICALC4 20 20"
Print #mFno, "PCFL ICALC5 20 20"
Print #mFno, "PCFL ICALG1 20 40"
Print #mFno, "PCFL ICALG2 20 40"
Print #mFno, "PCFL ICALG3 20 40"
Print #mFno, "PCFL ICALG4 20 40"
Print #mFno, "PCFL ICALG5 20 40"
Print #mFno, "PCFL ICALP1 2 7/2"
Print #mFno, "PCFL ICALP2 2 7/2"
Print #mFno, "PCFL ICALP3 2 7/2"
Print #mFno, "PCFL ICALP4 2 7/2"
Print #mFno, "PCFL ICALP5 2 7/2"
Print #mFno, "PCFL ICRETA 20 40"
Print #mFno, "PCFL ICCORA 20 40"
Print #mFno, "PCFL ICLMPC 2 7/2"
Print #mFno, "PCFL ICUMPC 2 7/2"
Close #mFno
If jj.Ovr_Wrt(cFx, True) Then Stop
If jj.Crt_Xls_FmFDF(cFx, cFfnFdf) Then Stop
Dim mWb As Workbook: If jj.Opn_Wb_RW(mWb, cFx, , True) Then Stop
End Function
Function Crt_Xls_FmHost_ForEdt(pNmtHost$, pDsn$, pSql$, pNPk As Byte, pFx$, Optional pLm$ = "") As Boolean
'Aim: create a {pFx} with NmWs {pNmtHost} by {pSql} through {pDns} with optional field name mapping in {pLm} with assume first {pNPK} is PK.
'{pFx} Fmt: Keep first NPK unchange, then a field [Changed], then list of pair of fields: [xx], [New xx]
Const cSub$ = "Crt_Xls_FmHost_ForEdt"
Const cNmt_DtaFmHost$ = "tmpCrtXls_FmHost_ForEdt_DtaFmHost"
Const cNmt_ForExp$ = "tmpCrtXls_FmHost_ForEdt_ForExp"
If jj.Dlt_Fil(pFx) Then ss.A 1: GoTo E
Dim mDteBeg As Date, mDteEnd As Date, mNRec&
If jj.Crt_Tbl_FmDSN_Sql(pDsn, pSql, cNmt_DtaFmHost, , mDteBeg, mDteEnd) Then ss.A 1: GoTo E
If jj.Crt_Tbl_ForEdtTbl(cNmt_DtaFmHost, pNPk, cNmt_ForExp) Then ss.A 2: GoTo E
If jj.Exp_Nmtq2Xls(cNmt_ForExp, pFx) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: Crt_Xls_FmHost_ForEdt = True: ss.B cSub, cMod, "pNmtHost,pDsn,pSql,pNPk,pFx,pLm", pNmtHost, pDsn, pSql, pNPk, pFx, pLm
End Function
Function Crt_Xls_FmHost_ForEdt_Tst() As Boolean
Const cSub$ = "Crt_Xls_FmHost_ForEdt_Tst"
Dim mNmtHost$, mDsn$, mSql$, mNPk As Byte, mFx$, mLm$
Dim mRslt As Boolean, mCase As Byte: mCase = 1
Select Case mCase
Case 1
    mNmtHost = "IIC"
    mDsn = "FETEST_ZBPCSF"
    mSql = "Select ICLAS,IID,ICDES,ICGL,ICCOGA,ICTAX,ICPPGL,ICALGL,ICALC1,ICALC2,ICALC3,ICALC4,ICALC5,ICALG1,ICALG2,ICALG3,ICALG4,ICALG5,ICALP1,ICALP2,ICALP3,ICALP4,ICALP5,ICRETA,ICCORA,ICLMPC,ICUMPC from IIC"
    mNPk = 1
    mFx = "c:\a.xls"
    mLm = ""
End Select
mRslt = jj.Crt_Xls_FmHost_ForEdt(mNmtHost, mDsn, mSql, mNPk, mFx, mLm)
jj.Shw_Dbg cSub, cMod, , "mNmtHost, mDsn, mSql,mNPK,mFx,mLm", mNmtHost, mDsn, mSql, mNPk, mFx, mLm
End Function
Function Crt_Xls_FmRs(pFxTar$, pRs As DAO.Recordset _
    , Optional pNmWs$ = "Data" _
    , Optional pRno& = 1 _
    , Optional pNoWsTit As Boolean = False) As Boolean
Const cSub$ = "Crt_Xls_FmRs"
'Aim: Create {pFxTar} having one worksheet of {pNmWs} from {pRs} @row {pRno}.  Note: pFxTar is always overwritten without ask
Dim mWb As Workbook, mWs As Worksheet
If jj.Crt_Wb(mWb, pFxTar, True, pNmWs) Then ss.A 2: GoTo E
Set mWs = mWb.Sheets(1)
Dim mR As Byte: If Not pNoWsTit Then mR = 1
pRs.MoveFirst
mWs.Range("A" & pRno + mR).CopyFromRecordset pRs
If Not pNoWsTit Then If jj.Set_WsTit_ByRs(mWs, pRs, pRno) Then ss.A 3: GoTo E
If jj.Cls_Wb(mWb, True) Then ss.A 4: GoTo E
Exit Function
R: ss.R
E: Crt_Xls_FmRs = True: ss.B cSub, cMod, "pFxTar,Rs,pNmWs,pRno,pNoWsTit", pFxTar, jj.ToStr_Rs(pRs), pNmWs, pRno, pNoWsTit
End Function
#If Tst Then
Function Crt_Xls_FmRs_Tst() As Boolean
Const cFx$ = "c:\tmp\bb.xls"
Dim mRs As DAO.Recordset: Set mRs = CurrentDb.OpenRecordset("Select * from mstBrand")
If jj.Crt_Xls_FmRs(cFx, mRs) Then Stop
Dim mWb As Workbook: If jj.Opn_Wb(mWb, cFx, , , True) Then Stop
End Function
#End If
Function Crt_Ws_FmRs(oWs As Worksheet, pWb As Workbook, pRs As DAO.Recordset _
    , Optional pNmWs$ = "Data" _
    , Optional pRno& = 1 _
    , Optional pNoWsTit As Boolean = False _
    ) As Boolean
Const cSub$ = "Crt_Ws_FmRs"
On Error GoTo R
If jj.Add_Ws(oWs, pWb, pNmWs) Then ss.A 1: GoTo E
Dim mR As Byte: If Not pNoWsTit Then mR = 1
pRs.MoveFirst
oWs.Range("A" & pRno + mR).CopyFromRecordset pRs
If Not pNoWsTit Then If jj.Set_WsTit_ByRs(oWs, pRs, pRno) Then ss.A 2: GoTo E
Exit Function
R: ss.R
E: Crt_Ws_FmRs = True: ss.B cSub, cMod
End Function
#If Tst Then
Function Crt_Ws_FmRs_Tst() As Boolean
Const cFx$ = "c:\tmp\bb.xls"
Dim mRs As DAO.Recordset: Set mRs = CurrentDb.OpenRecordset("Select * from mstBrand")
If jj.Crt_Xls_FmRs(cFx, mRs) Then Stop
Dim mWb As Workbook: If jj.Opn_Wb(mWb, cFx, , , True) Then Stop
Dim mWs As Worksheet: If jj.Crt_Ws_FmRs(mWs, mWb, mRs, "Data1") Then Stop
End Function
#End If
Function Crt_Fw(pFw$, Optional pOvrWrt As Boolean = False) As Boolean
Const cSub$ = "Crt_Fw"
Dim mWrd As Word.Document: If Crt_Wrd(mWrd, pFw, pOvrWrt) Then ss.A 1: GoTo E
If Cls_Wrd(mWrd, True) Then ss.A 2: GoTo E
Exit Function
E: Crt_Fw = True: ss.B cSub, cMod, "pFw,pOvrWrt", pFw, pOvrWrt
End Function
Function Crt_Wrd(oWrd As Word.Document, pFw$, Optional pOvrWrt As Boolean = False) As Boolean
Const cSub$ = "Crt_Wrd"
If pFw = "" Then ss.A 1, "pFw cannot be empty": GoTo E
If jj.Ovr_Wrt(pFw, pOvrWrt) Then ss.A 1: GoTo E
Set oWrd = gWrd.Documents.Add
On Error GoTo R
oWrd.SaveAs pFw
Exit Function
R: ss.R
E: Crt_Wrd = True: ss.B cSub, cMod, "pFw,pOvrWrt", pFw, pOvrWrt
End Function
Function Crt_Fp(pFp$, Optional pOvrWrt As Boolean = False) As Boolean
Const cSub$ = "Crt_Fp"
Dim mPpt As PowerPoint.Presentation: If Crt_Ppt(mPpt, pFp, pOvrWrt) Then ss.A 1: GoTo E
If Cls_Ppt(mPpt, True) Then ss.A 2: GoTo E
Exit Function
E: Crt_Fp = True: ss.B cSub, cMod, "pFp,pOvrWrt", pFp, pOvrWrt
End Function
Function Crt_Ppt(oPpt As PowerPoint.Presentation, pFp$, Optional pOvrWrt As Boolean = False) As Boolean
Const cSub$ = "Crt_Ppt"
If pFp = "" Then ss.A 1, "pFp cannot be empty": GoTo E
If jj.Ovr_Wrt(pFp, pOvrWrt) Then ss.A 1: GoTo E
Set oPpt = gPpt.Presentations.Add
On Error GoTo R
oPpt.SaveAs pFp
Exit Function
R: ss.R
E: Crt_Ppt = True: ss.B cSub, cMod, "pFp,pOvrWrt", pFp, pOvrWrt
End Function

