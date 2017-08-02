Attribute VB_Name = "xChk"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xChk"
'Function Chk_DocSml(pDocSml As MSXML2.DOMDocument60) As Boolean
'Const cSub$ = "Chk_DocSml"
'On Error GoTo R
'With pDocSml.parseError
'    If .errorCode <> 0 Then ss.A 1, "Error in pDocSml", , "Err,Lin,Pos,Txt,Url", .reason, .Line, .linepos, .srcText, .URL: GoTo E
'End With
'If pDocSml.ChildNodes(0).nodeName <> "Sml" Then ss.A 1, "Root Node Name must be Sml", , "Currnet Root Node Name", pDocSml.ChildNodes(0).nodeName: GoTo E
'Dim J%
'For J = 0 To pDocSml.ChildNodes(0).ChildNodes.Length - 1
'    If pDocSml.ChildNodes(0).ChildNodes(J).nodeName <> "Rec" Then ss.A 1, "The 2nd Lvl Node Name must be Rec", , "Currnet Node Name", pDocSml.ChildNodes(0).ChildNodes(J).nodeName: GoTo E
'Next
'Exit Function
'R: ss.R
'E: Chk_DocSml = True: ss.B cSub, cMod, ToStr_Doc(pDocSml)
'End Function
'#If Tst Then
'Function Chk_DocSml_Tst() As Boolean
'Dim mDocSml As New DOMDocument60
'mDocSml.loadXML "<Sml><Rec><aa>lsdkjf</aa></Rec></Sml>"
'If Chk_DocSml(mDocSml) Then Stop
'End Function
'#End If
Function Chk_No2Par(pNmt$, pNmFldChd$, pNmFldPar$) As Boolean
'Aim: Chk if {pNmFldChd} in {pNmt} has 2 or more parent.
Const cSub$ = "Chk_No2Par"
Dim mNmt$: mNmt = jj.Rmv_SqBkt(pNmt)
Dim mNmtTmp$: mNmtTmp = jj.Fmt_Str("[#Chk_No2Par_{2}_{0}_{1}]", pNmFldChd, pNmFldPar, mNmt)
Dim mSql$: mSql = jj.Fmt_Str("Select Distinct {0},{1} into {3} from [{2}]", pNmFldChd, pNmFldPar, mNmt, mNmtTmp)
If jj.Run_Sql(mSql) Then ss.A 1: GoTo E
mSql = jj.Fmt_Str("Select Distinct {0},Count(*) from {1} group by {0} having Count(*)>1", pNmFldChd, mNmtTmp)
Dim mA$: If jj.Fnd_Lv_FmDistSql(mA, mSql) Then ss.A 2: GoTo E
If mA <> "" Then ss.A 3, "The list of value are child having 2 or more parent", , "The List", mA: GoTo E
If jj.Dlt_Tbl(mNmtTmp) Then ss.A 4: GoTo E
Exit Function
R: ss.R
E: Chk_No2Par = True: ss.B cSub, cMod, "pNmt,pNmFldChd,pNmFldPar", pNmt, pNmFldChd, pNmFldPar
End Function
#If Tst Then
Function Chk_No2Par_Tst() As Boolean
Const cNmt$ = "[#Tmp]"
If jj.Crt_Tbl_FmLoFld(cNmt, "aa int, bb int") Then Stop
If jj.Run_Sql("INsert into " & cNmt & " (aa,bb) values (1,1)") Then Stop
If jj.Run_Sql("INsert into " & cNmt & " (aa,bb) values (1,2)") Then Stop
If jj.Run_Sql("INsert into " & cNmt & " (aa,bb) values (1,3)") Then Stop
If jj.Run_Sql("INsert into " & cNmt & " (aa,bb) values (1,4)") Then Stop
If jj.Run_Sql("INsert into " & cNmt & " (aa,bb) values (2,5)") Then Stop
If jj.Run_Sql("INsert into " & cNmt & " (aa,bb) values (2,6)") Then Stop
Debug.Print "Should be False -->"; jj.Chk_No2Par(cNmt, "bb", "aa")
If jj.Run_Sql("INsert into " & cNmt & " (aa,bb) values (2,6)") Then Stop
Debug.Print "Should be False -->"; jj.Chk_No2Par(cNmt, "bb", "aa")
If jj.Run_Sql("INsert into " & cNmt & " (aa,bb) values (1,6)") Then Stop
Debug.Print "Should be True -->"; jj.Chk_No2Par(cNmt, "bb", "aa")
If jj.Run_Sql("INsert into " & cNmt & " (aa,bb) values (1,Null)") Then Stop
If jj.Run_Sql("INsert into " & cNmt & " (aa,bb) values (2,Null)") Then Stop
Debug.Print "Should be True -->"; jj.Chk_No2Par(cNmt, "bb", "aa")
End Function
#End If
Function Chk_LgcVer(oIsOk As Boolean, pFbTar$, pFbSrc$) As Boolean
Const cSub$ = "Chk_LgcVer"
'Aim: Return error {pFbTar} is an older version than {pFbSrc} by comparing the tblVer->Ver in both Ffn.  If no tblVer, it will be considered as older.  If both does not have tblVer, it will be considered as error
If VBA.Dir(pFbTar) = "" Then ss.A 1, "pFbTar not exist", "pFbTar", pFbTar: GoTo E
If VBA.Dir(pFbSrc) = "" Then ss.A 2, "pFbSrc not exist", "pFbSrc", pFbSrc: GoTo E
Dim mVerTar, mVerSrc
oIsOk = False
On Error GoTo R
jj.Set_Silent
If jj.Fnd_ValFmSql(mVerSrc, jj.Q_S(pFbSrc, "Select Ver from tblVer in '*'")) Then
    Dim mDbSrc As DAO.Database: If jj.Opn_Db_RW(mDbSrc, pFbSrc) Then ss.A 3: GoTo E
    If jj.Crt_TblVer(mDbSrc) Then ss.A 4, "Cannot create TblVer": GoTo E
    GoTo X
End If
If jj.Fnd_ValFmSql(mVerTar, jj.Q_S(pFbTar, "Select Ver from tblVer in '*'")) Then ss.A 5, "Tar Mdb has no tblVer": GoTo E
If mVerTar > mVerSrc Then ss.A 1, "Impossible: Tar's ver is newer", eImpossibleReachHere: GoTo E
If mVerTar = mVerSrc Then oIsOk = True
GoTo X
R: ss.R
E: Chk_LgcVer = True: ss.B cSub, cMod, "pFbTar$, pFbSrc$", pFbTar$, pFbSrc$
X: jj.Set_Silent_Rst
   jj.Cls_Db mDbSrc
End Function
#If Tst Then
Function Chk_LgcVer_Tst() As Boolean
Dim mIsOk As Boolean
Const cFm$ = "c:\Tmp\aa.mdb"
Const cTo$ = "c:\Tmp\bb.mdb"
If jj.Crt_Fb(cFm, True) Then Stop
If jj.Crt_Fb(cTo, True) Then Stop
If jj.Chk_LgcVer(mIsOk, cTo, cFm) Then Stop
Debug.Print mIsOk
End Function
#End If
Function Chk_Struct_Nmtq(pNmtq$, pLnFld$, pDb As DAO.Database) As Boolean
Const cSub$ = "Chk_Nmtq"
On Error GoTo R
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim mFlds As DAO.Fields: Set mFlds = mDb.QueryDefs(pNmtq).Fields
Chk_Struct_Nmtq = jj.Chk_Struct_Flds(mFlds, pLnFld)
R: ss.R
E: Chk_Struct_Nmtq = True: ss.B cSub, cMod, "pNmtq,pLnFld", pNmtq, pLnFld
End Function
Function Chk_Struct_Flds(pFlds As DAO.Fields, pLnFld$) As Boolean
Const cSub$ = "Chk_Flds"
On Error GoTo R
Chk_Struct_Flds = pLnFld <> jj.ToStr_Flds(pFlds)
Exit Function
R: ss.R
E: Chk_Struct_Flds = True: ss.B cSub, cMod, "pFlds,pLnFld", jj.ToStr_Flds(pFlds), pLnFld
End Function
Function Chk_Struct_Rs(pRs As DAO.Recordset, pLnFld$) As Boolean
Const cSub$ = "Chk_Struct_Rs"
On Error GoTo R
Chk_Struct_Rs = pLnFld <> jj.ToStr_Flds(pRs.Fields)
Exit Function
R: ss.R
E: Chk_Struct_Rs = True: ss.B cSub, cMod, "pRs,pLnFld", jj.ToStr_Rs_NmFld(pRs), pLnFld
End Function
Function Chk_Struct_Nmq(pNmq$, pLnFld$, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Chk_Struct_Nmq"
On Error GoTo R
Chk_Struct_Nmq = jj.Chk_Struct_Flds(jj.Cv_Db(pDb).QueryDefs(jj.Rmv_SqBkt(pNmq)).Fields, pLnFld)
Exit Function
R: ss.R
E: Chk_Struct_Nmq = True: ss.B cSub, cMod, "pNmq,pLnFld,pDb", pNmq, jj.ToStr_Db(pDb), pLnFld
End Function
Function Chk_Struct_Tbl(pNmt$, pLnFld$, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Chk_Struct_Tbl"
On Error GoTo R
Chk_Struct_Tbl = Chk_Struct_Flds(jj.Cv_Db(pDb).TableDefs(jj.Rmv_SqBkt(pNmt)).Fields, pLnFld)
Exit Function
R: ss.R
E: Chk_Struct_Tbl = True: ss.B cSub, cMod, "pNmt,pLnFld,pDb", pNmt, jj.ToStr_Db(pDb), pLnFld
End Function
#If Tst Then
Function Chk_Struct_Tbl_Tst() As Boolean
If jj.Chk_Struct_Tbl("mstBrand", "BrandId,Brand,ItemClass_FG,ItemClass_Cmp,IsAlwSel,IsCurr1") Then Stop
End Function
#End If
Function Chk_Struct_Tbl_SubSet(pNmt$, pLnFld$, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Chk_Struct_Tbl_SubSet"
On Error GoTo R
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
If Not jj.IsTbl(pNmt, mDb) Then ss.A 1, "Given pNmt not exit in pDb": GoTo E
Dim mLnFld$: mLnFld = jj.ToStr_Flds(mDb.TableDefs(pNmt).Fields)
If Not jj.IsSubSet_Ln(mLnFld, pLnFld) Then ss.A 2, "Given pLnFld is not a subset of the given table structure", , "Table Structure of pNmt", mLnFld: GoTo E
Exit Function
R: ss.R
E: Chk_Struct_Tbl_SubSet = True: ss.B cSub, cMod, "pNmt,pLnFld,pDb", pNmt, pLnFld, jj.ToStr_Db(pDb), pLnFld
End Function
Function Chk_Cell(pWs As Worksheet, pAdr$, pVal$) As Boolean
Const cSub$ = "Chk_Cell"
If pWs.Range(pAdr).Value = pVal Then Exit Function
ss.A 1, "Expected value not found in ws", eUsrInfo
GoTo E
E: Chk_Cell = True: ss.B cSub, cMod, "Ws, Addr, Value", pWs.Name, pAdr, pVal
End Function
Function Chk_Cell_InRge(oRge As Range, pRge As Range, pVal$) As Boolean
'Aim Check with cell in {pRge} contains given {pVal} and return the cell address in {oCellAdr}
Const cSub$ = "Chk_Cell_InRge"
For Each oRge In pRge.Cells
    If oRge.Value = pVal Then Exit Function
Next
ss.A 1, "Expected value not found in ws", eUsrInfo: GoTo E
GoTo E
E: Chk_Cell_InRge = True: ss.B cSub, cMod, "Ws Nam, pRge, Value", pRge.Worksheet.Name, pRge.Address, pVal
End Function
Function Chk_DupKey(pNmt$, pNPk%) As Boolean
'Aim: Chk if first {pNPK} fields in {pNmt} has duplicate.  Return true is there is duplicate
Const cSub$ = "Chk_DupKey"
'-Start
Dim mLst$: mLst = jj.ToStr_Nmt(pNmt, , , pNPk - 1)
Dim mSql$: mSql = jj.Fmt_Str("Select Distinct {0}, Count(*) as Cnt from {1} group by {0} having Count(*)>1", mLst, pNmt)
With CurrentDb.OpenRecordset(mSql)
    If .AbsolutePosition = -1 Then .Close: Exit Function
    .Close
End With
Chk_DupKey = True
End Function
Function Chk_DupKey_Tst() As Boolean
Debug.Print "DupKey(mstAllBrand,1) = " & jj.Chk_DupKey("mstAllBrand", 1)
Debug.Print "DupKey(mstAllBrand,2) = " & jj.Chk_DupKey("mstAllBrand", 2)
End Function
Function Chk_Host_ByFrm(oHostSts As eHostSts, _
    pNmtHost$, pDsn$, pFrm As Access.Form, pLmPk$, pLnFld$ _
    ) As Boolean
'Aim: Check host has one same record as local else return Er.
'     Use {pLmPk} & {pFrm}'s .OldValue to build {oPKCndn} to get Rs from {pNmtHost} through {pDsn}.  The Qry is qryChkHostByFrm
'     {oHostSts} {oPKCndn} {oAn2V_PK} {oAn2V_Lst} {oAyTypSim_Lst} will be returned
'     If Host & Form's record are different, host's record will be copied to form's controls' value
'{oHostSts}:
'   eUnExpectedErr
'   e0Rec
'   e1Rec
'   e2Rec
'   eHostCpyToFrm
Const cSub$ = "Chk_Host_ByFrm"
Const cNmq = "qryChkHostByFrm"
oHostSts = eUnExpectedErr

'Fnd Host mRs by pFrm,pLmPk,pLnFld + pNmtHost,pDsn
Dim mRs As DAO.Recordset
Do
    Dim mPKCndn$: If jj.Bld_LExpr_InFrm(mPKCndn, pFrm, pLmPk) Then ss.A 2: GoTo E
    'Get mRs from host by setting {cNmq}
    Dim mAn_Frm$(), mAn_Host$(): If jj.Brk_Lm_To2Ay(mAn_Frm, mAn_Host, pLnFld) Then ss.A 3: GoTo E
    Dim mSql$: mSql = jj.ToSql_Sel(pNmtHost, Join(mAn_Host, cComma), mPKCndn)
    If jj.Crt_Qry_ByDSN(cNmq, mSql, pDsn, True) Then ss.A 4: GoTo E
    If jj.Opn_Rs_ByNmq(mRs, cNmq) Then ss.A 5: GoTo E
Loop Until True

'Check if No Rec In Host ---------
With mRs
    If .AbsolutePosition = -1 Then
        .Close
        ss.A 6, "There is not corresponding record in Host."
        oHostSts = e0Rec
        GoTo E
    End If

'Check if 2orMore Rec in Host ----
    .MoveNext
    If Not .EOF Then
        .Close
        ss.A 7, "There is 2 or more corresponding record in Host|That host has data error": GoTo E
        oHostSts = e2Rec
        GoTo E
    End If
    .MoveFirst

'Compare
    Dim mIsSam As Boolean: If jj.Cmp_Rs_VsFrm(mIsSam, mRs, pFrm, pLnFld) Then mRs.Close: ss.A 7: GoTo E
    If Not mIsSam Then
        ''If not same, Cpy from Host's mRs to pFrm's controls' value & return
        If jj.Cpy_Rs_ToFrm(mRs, pFrm, pLnFld) Then ss.A 3: GoTo E
        .Close
        ss.A 8, "There is value different between current record and host data|Current Record is copied from host data."
        oHostSts = eHostCpyToFrm
        GoTo E
    End If
    
    .Close
    oHostSts = e1Rec
End With
Exit Function
R: ss.R
E: Chk_Host_ByFrm = True: ss.B cSub, cMod, "pNmtHost, mPKCndn", pNmtHost, mPKCndn
End Function
#If Tst Then
Function Chk_Host_ByFrm_Tst() As Boolean
Const cSub$ = "Chk_Host_ByFrm_Tst"
Dim cNmtHost$:  cNmtHost$ = "IIC"
Dim cDsn$:      cDsn$ = "FETEST_ZBPCSF"
Dim cLstPK$:    cLstPK$ = "ICLAS"
Dim cLnFld$:    cLnFld$ = "ICDES"
Const cNmFrm$ = "frmIIC_Tst"
Dim mFrm As Access.Form: If jj.Opn_Frm(cNmFrm, , , mFrm) Then Stop: GoTo E
mFrm.Visible = False
Dim mHostSts As eHostSts
jj.Shw_Dbg cSub, cMod
Dim mPKCndn$, mAyNm2V_PK() As tNm2V, mAyNm2V_Lst() As tNm2V, mAyTypSim_Lst() As eTypSim, mRslt As Boolean
Dim mCase As Byte: mCase = 2
For mCase = 1 To 3
    Debug.Print "Case" & mCase & "--------------------------------------------"
    Select Case mCase
    Case 1
    Case 2
        mFrm.Recordset.MoveNext
    Case 3
        mFrm.Recordset.MoveNext
    End Select
    mRslt = jj.Chk_Host_ByFrm(mHostSts, cNmtHost, cDsn, mFrm, cLstPK, cLnFld)
    Debug.Print jj.ToStr_LpAp(vbLf, "mRslt, mHostSts, cNmtHost, cDsn, cLstPK, cLnFld", mRslt, jj.ToStr_HostSts(mHostSts), cNmtHost, cDsn, cLstPK, cLnFld)
    Debug.Print "---------"
    Debug.Print "mAn2_PK"
    Debug.Print jj.ToStr_An2V(mAyNm2V_PK)
    Debug.Print "----------"
    Debug.Print "mAn2_Lst"
    Debug.Print jj.ToStr_An2V(mAyNm2V_Lst)
    Debug.Print
Next
Exit Function
E: Chk_Host_ByFrm_Tst = True
End Function
#End If
