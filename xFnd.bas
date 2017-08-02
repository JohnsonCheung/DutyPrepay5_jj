Attribute VB_Name = "xFnd"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xFnd"
Function Fnd_MsgBoxSty(pTypMsg As eTypMsg) As VbMsgBoxStyle
Dim mA As VbMsgBoxStyle
Select Case pTypMsg
    Case eTypMsg.eCritical, eTypMsg.ePrmErr: mA = vbCritical
    Case eTypMsg.eWarning: mA = vbExclamation
    Case eTypMsg.eTrc, eTypMsg.eUsrInfo: mA = vbInformation
    Case Else: mA = vbInformation
End Select
mA = mA Or vbDefaultButton1
If jj.SysCfg_IsDbg Then Fnd_MsgBoxSty = mA Or vbYesNo
End Function
Function Fnd_CnoLas(oCnoLas As Byte, pRge As Range) As Boolean
On Error GoTo R
Dim mRge As Range: Set mRge = pRge(1, 256 - pRge.Column)
Shw_AllCols pRge.Parent
oCnoLas = mRge.End(xlToLeft).Column
Exit Function
R: Fnd_CnoLas = True
End Function
Function Fnd_Qt(oQt As QueryTable, pWs As Worksheet, pNmQt$) As Boolean
On Error GoTo R
Set oQt = pWs.QueryTables(pNmQt)
Exit Function
R: Fnd_Qt = True
End Function

Function Fnd_NxtFfn$(pFfn$)
'Aim: If pFfn exist, find next Ffn by adding (n) to the end of the file name.
Const cSub$ = "Ffn_NxtFfn"
If VBA.Dir(pFfn) = "" Then Fnd_NxtFfn = pFfn: Exit Function
Dim mP%: mP = InStrRev(pFfn, ".")
Dim mA$, mB$
If mP = 0 Then
    mA = pFfn
Else
    mA = Left(pFfn, mP - 1)
    mB = mID(pFfn, mP)
End If
Dim J%
For J = 0 To 100
    Dim mFfn$: mFfn = mA & "(" & J & ")" & mB
    If VBA.Dir(mFfn) = "" Then Fnd_NxtFfn = mFfn: Exit Function
Next
ss.A 1, "Quit impossible to reach here.. Having 100 next file exist"
E: ss.B cSub, cMod, "pFfn", pFfn
End Function
Function Fnd_AyCnoImpFld(oAyCno() As Byte, oAmFld() As tMap, pRge As Range _
    , Optional pRithNmFld% = -1, Optional pRithImp% = -4) As Boolean
'Aim: Whenever @ {pRithImp} & {pRithNmFld}, there is value a TypFld (vbString) & NmFld (vbString),
'     The column is a Import Field.  Put its name, type & Cno into {oAnFld}, {oAyTypFld} & {oAyCno}
Const cSub$ = "Fnd_AyCnoImpFld"
On Error GoTo R
jj.Clr_Am oAmFld
jj.Clr_AyByt oAyCno
Dim mRnoImp&: mRnoImp = pRge.Row + pRithImp
Dim mRnoNmFld&: mRnoNmFld = pRge.Row + pRithNmFld
Dim mCnoBeg As Byte: mCnoBeg = pRge.Column

Dim mWs As Worksheet: Set mWs = pRge.Parent
Dim mV: mV = mWs.Cells(mRnoImp, mCnoBeg).Value
If VarType(mV) <> vbString Then ss.A 1, jj.Fmt_Str("Cell({0},{1}) must be vbString", mRnoImp, mCnoBeg), , "The Cell", mV: GoTo E
If mV <> "Import:" & mWs.Name Then ss.A 1, jj.Fmt_Str("Cell({0},{1}) must be 'Import:{2}'", mRnoImp, mCnoBeg, mWs.Name), , "The Cell", mV: GoTo E
Dim iCno As Byte, N As Byte, mCnoLas As Byte: If jj.Fnd_CnoLas(mCnoLas, pRge(0, 1)) Then ss.A 1: GoTo E
If mCnoLas - pRge.Column < 1 Then ss.A 2, "There should at least 2 columns Id & Nam", , "pRge.Column,mCnoLas", pRge.Column, mCnoLas: GoTo E
For iCno = mCnoBeg To mCnoLas
    mV = mWs.Cells(mRnoNmFld, iCno).Value
    Dim mT: mT = mWs.Cells(mRnoImp, iCno).Value
    If VarType(mV) = vbString Then
        Select Case VarType(mT)
        Case vbString: ReDim Preserve oAmFld(N), oAyCno(N)
                       oAmFld(N).F1 = mV: oAyCno(N) = iCno: oAmFld(N).F2 = mT
                       N = N + 1
        Case vbEmpty
        Case Else:     ss.A 2, "It has a vbString field name, but a invalid Type", , "iCno,NmFld,TypFld", iCno, mV, mT
        End Select
    Else
        If VarType(mT) <> vbEmpty Then ss.A 3, "A non field name cannot have non-empty Type", , "iCno,non-empty-type", iCno, mT
    End If
Next
If InStr(oAmFld(0).F1, "_") > 0 Then
    oAmFld(0).F2 = "Text 255"
Else
    oAmFld(0).F2 = "Long"
End If
Exit Function
R: ss.R
E: Fnd_AyCnoImpFld = True: ss.B cSub, cMod, "pRge,pRithImp", jj.ToStr_Rge(pRge), pRithImp
End Function
#If Tst Then
Function Fnd_AyCnoImpFld_Tst() As Boolean
Dim mWb As Workbook: If jj.Crt_Wb(mWb, "c:\tmp\bb.xls", True, "Sheet1") Then Stop: GoTo E

'^^
mWb.Sheets(1).Cells(1, 1).Value = "Import:Sheet1"

mWb.Sheets(1).Cells(4, 1).Value = "Id_Id"

mWb.Sheets(1).Cells(1, 2).Value = "Text 1"
mWb.Sheets(1).Cells(4, 2).Value = "Nm1"

mWb.Sheets(1).Cells(1, 5).Value = "Text 2"
mWb.Sheets(1).Cells(4, 5).Value = "Nm2"

mWb.Sheets(1).Cells(1, 6).Value = "Text 3"
mWb.Sheets(1).Cells(4, 6).Value = "Nm3"

mWb.Sheets(1).Cells(1, 7).Value = "Text 4"
mWb.Sheets(1).Cells(4, 7).Value = "Nm4"


Dim mAmFld() As tMap, mAyCno() As Byte: If jj.Fnd_AyCnoImpFld(mAyCno, mAmFld, mWb.Sheets(1).Range("A5")) Then Stop
Debug.Print jj.ToStr_LpAp(cComma, "mAyCno", ToStr_AyByt(mAyCno))
Debug.Print jj.ToStr_LpAp(cComma, "mAmFld", ToStr_Am(mAmFld, " "))
jj.Shw_DbgWin
mWb.Application.Visible = True
Stop
jj.Cls_Wb mWb, False, True
E: Fnd_AyCnoImpFld_Tst = True
X: jj.Cls_Wb mWb, , True
End Function
#End If
Function Fnd_AyRecCnt(oAyRecCnt&(), pAnt$(), Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: Find {oAyRecCnt} of each table of {pDb!pLnt}
Const cSub$ = "Fnd_AyRecCnt"
On Error GoTo R
Dim N%, J%
For J = 0 To jj.Siz_Ay(pAnt) - 1
    ReDim Preserve oAyRecCnt(N)
    oAyRecCnt(N) = Fct.RecCnt(pAnt(J), pDb)
    N = N + 1
Next
If N = 0 Then jj.Clr_AyLng oAyRecCnt
Exit Function
R: ss.R
E: Fnd_AyRecCnt = True: ss.B cSub, cMod, "pAnt,pDb", jj.ToStr_Ays(pAnt), jj.ToStr_Db(pDb)
End Function
Function Fnd_AyCnoColr(oAyCno() As Byte, oAyColr&(), pRge As Range, pRnoColrIdx&) As Boolean
'Aim: Find the color of a row {pRnoColrIdx} into {oAyCno} & {oAyColr}.  Row pRge(0,1) will be used to detect the start and end column
Const cSub$ = "Fnd_AyCnoColr"
On Error GoTo R
jj.Clr_AyLng oAyColr
jj.Clr_AyByt oAyCno
Dim iCno As Byte, N%, mCnoLas As Byte: If jj.Fnd_CnoLas(mCnoLas, pRge(0, 1)) Then ss.A 1: GoTo E
Dim mWs As Worksheet: Set mWs = pRge.Parent
For iCno = pRge.Column To mCnoLas
    Dim mRge As Range: Set mRge = mWs.Cells(pRnoColrIdx, iCno)
    Dim mColr&: mColr = mRge.Interior.Color
    If mColr <> jj.g.cColrNo Then
        ReDim Preserve oAyCno(N), oAyColr(N)
        oAyCno(N) = iCno
        oAyColr(N) = mRge.Interior.Color
        N = N + 1
    End If
Next
Exit Function
R: ss.R
E: Fnd_AyCnoColr = True: ss.B cSub, cMod, "pRge,pRnoColrIdx", jj.ToStr_Rge(pRge), pRnoColrIdx
End Function
#If Tst Then
Function Fnd_AyCnoColr_Tst() As Boolean
Dim mWb As Workbook: If jj.Crt_Wb(mWb, "c:\tmp\aa.xls", True, "Sheet1") Then Stop: GoTo E
Dim mAyCno() As Byte, mAyColr&()
mWb.Sheets(1).Cells(2, 5).Interior.Color = 123
If jj.Fnd_AyCnoColr(mAyCno, mAyColr, mWb.Sheets(1), 2) Then Stop
Debug.Print ToStr_AyByt(mAyCno)
Debug.Print ToStr_AyLng(mAyColr)
jj.Shw_DbgWin
Stop
jj.Cls_Wb mWb, False, True
E: Fnd_AyCnoColr_Tst = True
End Function
#End If
Function Fnd_AyRno_Visible(oAyRno&(), pRge As Range) As Boolean
'Aim: Find oAyRno() which is not hidden row started at {pRge}
Const cSub$ = "Fnd_AyRno_Visible"
On Error GoTo R
jj.Clr_AyLng oAyRno
Dim iRno&, N%, mWs As Worksheet: Set mWs = pRge.Parent
For iRno = pRge.Row To pRge(1, 1).End(xlDown).Row
    If Not mWs.Rows(iRno).Hidden Then
        ReDim Preserve oAyRno(N): oAyRno(N) = iRno: N = N + 1
    End If
Next
Exit Function
R: ss.R
E: Fnd_AyRno_Visible = True: ss.B cSub, cMod, "pRge,pCno", jj.ToStr_Rge(pRge)
End Function
'Function Fnd_DocSml_ByRno(oDocSml As DOMDocument60, pWs As Worksheet, pRno&) As Boolean
''Aim: find {oLnv}, which is a string of one line one Name=Value from {pRno} of {pWs} by using all Ws Names begins with x in {pWs}
'jj.Clr_Doc oDocSml
'Dim N%, J%, mNod As MSXML2.IXMLDOMNode, mChd As IXMLDOMNode
'Set mChd = oDocSml.createNode(NODE_ELEMENT, "SML", ""): Set mNod = oDocSml.appendChild(mChd)
'Set mChd = oDocSml.createNode(NODE_ELEMENT, "Rec", ""): Set mNod = oDocSml.appendChild(mChd)
'For J = 0 To pWs.Names.Count - 1
'    With pWs.Names(J)
'        If Left(.Name, Len(pWs.Name) + 2) = pWs.Name & "!x" Then
'            Dim mNm$: mNm = mID(.Name, Len(pWs.Name) + 3)
'            Set mChd = oDocSml.createNode(NODE_ELEMENT, mNm, "")
'            mChd.Text = Nz(pWs.Cells(pRno, .RefersToRange.Column).Value, "")
'            mNod.appendChild mChd
'        End If
'    End With
'Next
'End Function
'Function Fnd_DocSml_ByFfn(oDocSml As MSXML2.DOMDocument60, pFfn$) As Boolean
''Aim: Read content from {pFfn} and create {oDocSml}
'Const cSub$ = "Fnd_DocSml_ByFfn"
'jj.Clr_Doc oDocSml
'oDocSml.Load pFfn
'If jj.Chk_DocSml(oDocSml) Then ss.A 1: GoTo E
'Exit Function
'E: Fnd_DocSml_ByFfn = True: ss.B cSub, cMod, jj.ToStr_Doc(oDocSml)
'End Function
Function Fnd_Anq_wSubStr(oAnq$(), pSubStr$, Optional pDb As DAO.Database = Nothing, Optional pSilent) As Boolean
Const cSub$ = "Fnd_Anq_wSubStr"
On Error GoTo R
Dim iQry As QueryDef
For Each iQry In CurrentDb.QueryDefs
    If Left(iQry.Name, 1) <> "~" Then
        If InStr(iQry.Sql, pSubStr) > 0 Then jj.Add_AyEle oAnq, iQry.Name
    End If
Next
Exit Function
R: ss.R
E: Fnd_Anq_wSubStr = True: If Not pSilent Then ss.B cSub, cMod, "pStr"
End Function
Function Fnd_AnFld_ReqTxt(oAnFld$(), pNmt$, Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: Find {oAnFld} of {pNmt} which is either (Text or Memo) is IsReq
Const cSub$ = "Fnd_AnFld_ReqTxt"
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim iFld As DAO.Field
jj.Clr_Ays oAnFld
Dim N%: N = 0
For Each iFld In mDb.TableDefs(pNmt).Fields
    Select Case iFld.Type
    Case DAO.DataTypeEnum.dbText, DAO.DataTypeEnum.dbMemo
        If iFld.Required Then ReDim Preserve oAnFld(N): oAnFld(N) = iFld.Name: N = N + 1
    End Select
Next
Exit Function
R: ss.R
E: Fnd_AnFld_ReqTxt = True: ss.B cSub, cMod, "pNmt", pNmt
End Function
#If Tst Then
Function Fnd_AnFld_ReqTxt_Tst() As Boolean
Dim mDb As DAO.Database: If jj.Opn_Db_R(mDb, "p:\workingdir\MetaAll.mdb") Then Stop: GoTo E
Dim mAnt$(): If jj.Fnd_AyVFmSql(mAnt, "Select NmTbl from [$Tbl] where NmTbl=PKey", mDb) Then Stop: GoTo E
Dim J%, N%: N = jj.Siz_Ay(mAnt)
For J = 0 To N - 1
    Dim mAnFld$(): If jj.Fnd_AnFld_ReqTxt(mAnFld, "$" & mAnt(J), mDb) Then Stop: GoTo E
    Debug.Print mAnt(J) & ":" & jj.ToStr_Ays(mAnFld)
Next
Exit Function
E: Fnd_AnFld_ReqTxt_Tst = True
End Function
#End If
Function Fnd_Nm_ById(oNm$, pItm$, pId&) As Boolean
'Aim: Assume there is a table [${pItm}] having 2 fields: [{pItm}] & [Nm{pItm}].  Use [{pItm}] to find the name in table.
Const cSub$ = "Fnd_Nm_ById"
If jj.Fnd_ValFmSql(oNm, jj.Fmt_Str("Select Nm{0} from [${0}] where {0}={1}", pItm, pId)) Then ss.A 1, "pId not find in [${pItm}]": GoTo E
Exit Function
E: Fnd_Nm_ById = True: ss.B cSub, cMod, "pItm,pId", pItm, pId
End Function
Function Fnd_Id_ByNm(oId&, pItm$, pNm$) As Boolean
'Aim: Assume there is a table [${pItm}] having 2 fields: [{pItm}] & [Nm{pItm}].  Use [Nm{pItm}] to find the Id in table.
Const cSub$ = "Fnd_Id_ByNm"
If jj.Fnd_ValFmSql(oId, jj.Fmt_Str("Select {0} from [${0}] where Nm{0}='{1}'", pItm, pNm)) Then ss.A 1, "pNm not find in [${pItm}]": GoTo E
Exit Function
E: Fnd_Id_ByNm = True: ss.B cSub, cMod, "pItm,pNm", pItm, pNm
End Function
Function Fnd_AyRoot(oAyRoot&(), pNmt$, pPar$, pChd$) As Boolean
'Aim: In {pNmt} fields {pPar} & {pChr} are parent & child relation.  It is to find all Id of root to {oAyTblRoot}
Const cSub$ = "Fnd_AyRoot"
On Error GoTo R
Dim mNmtChd$: mNmtChd$ = "[##CHD" & Format(Now, "YYYYMMDD HHMMSS") & "]"
Dim mNmtPar$: mNmtPar$ = "[##PAR" & Format(Now, "YYYYMMDD HHMMSS") & "]"
Dim mNmt$: mNmt = jj.Q_S(pNmt, "[]")
Dim mSql$
mSql = jj.Fmt_Str("Select Distinct {0} into {1} from {2}", pPar, mNmtPar, mNmt): If jj.Run_Sql(mSql) Then ss.A 1: GoTo E
mSql = jj.Fmt_Str("Select Distinct {0} into {1} from {2}", pChd, mNmtChd, mNmt): If jj.Run_Sql(mSql) Then ss.A 2: GoTo E
mSql = jj.Fmt_Str_ByLpAp("Select {pPar} from {mNmtPar} p left join {mNmtChd} c on p.{pPar}=c.{pChd} where c.{pChd} is null order by {pPar}" _
    , "pPar,pChd,mNmtPar,mNmtChd,mNmt", pPar, pChd, mNmtPar, mNmtChd, mNmt)
Fnd_AyRoot = jj.Fnd_AyVFmSql(oAyRoot, mSql)
jj.Run_Sql "Drop Table" & mNmtChd
jj.Run_Sql "Drop Table" & mNmtPar
Exit Function
R: ss.R
E: Fnd_AyRoot = True: ss.B cSub, cMod, "pNmt,pPar,pChd", pNmt, pPar, pChd
End Function
#If Tst Then
Function Fnd_AyRoot_Tst() As Boolean
If jj.Crt_Tbl_FmLnkLnt("p:\workingdir\MetaAll.mdb", "$TblR,$Tbl") Then Stop: GoTo E
Dim mAyRoot&(): If jj.Fnd_AyRoot(mAyRoot, "$TblR", "Tbl", "TblTo") Then Stop: GoTo E
Dim mAnt$(): If jj.Fnd_AyVFmSql(mAnt, "Select NmTbl from [$Tbl] where Tbl in (" & jj.ToStr_AyLng(mAyRoot) & ")") Then Stop: GoTo E
Debug.Print jj.ToStr_Ays(mAnt)
Debug.Print jj.ToStr_AyLng(mAyRoot)
Exit Function
E: Fnd_AyRoot_Tst = True
End Function
#End If
Function Fnd_AyChd_ByRoot(oAyId&(), pNmt$, pRoot&, pPar$, pChd$ _
    , Optional pRootFirst As Boolean = False _
    , Optional pKeepAyId As Boolean = False) As Boolean
'Aim: In {pNmt} fields {pPar} & {pChr} are parent & child relation.  It is to find all Id of the tree of {pRoot} into {oAyId}
Const cSub$ = "Fnd_AyChd_ByRoot"
On Error GoTo R
If Not pKeepAyId Then jj.Clr_AyLng oAyId
On Error GoTo R
Dim N%, J
If pRootFirst Then If jj.Add_AyEleLng(oAyId, pRoot, True) Then Exit Function
Dim mSql$: mSql = jj.Fmt_Str_ByLpAp("Select {pChd} from {pNmt} where {pPar}={pRoot}", "pChd,pNmt,pPar,pRoot", pChd, jj.Q_S(pNmt, "[]"), pPar, pRoot)
Dim mAyId&(): If jj.Fnd_AyVFmSql(mAyId, mSql) Then ss.A 2: GoTo E
For J = 0 To jj.Siz_Ay(mAyId) - 1
    If jj.Fnd_AyChd_ByRoot(oAyId, pNmt, mAyId(J), pPar, pChd, pRootFirst, True) Then ss.A 3: GoTo E
Next
If Not pRootFirst Then If jj.Add_AyEleLng(oAyId, pRoot, True) Then Exit Function
Exit Function
R: ss.R
E: Fnd_AyChd_ByRoot = True: ss.B cSub, cMod, "pNmt,pRoot,pPar,pChd,pRootFirst,pKeepAyId", pNmt, pRoot, pPar, pChd, pRootFirst, pKeepAyId
End Function
#If Tst Then
Function Fnd_AyChd_ByRoot_Tst() As Boolean
If jj.Crt_Tbl_FmLnkLnt("p:\workingdir\MetaAll.mdb", "$TblR,$Tbl") Then Stop: GoTo E
Dim mAyRoot&(): If jj.Fnd_AyRoot(mAyRoot, "$TblR", "Tbl", "TblTo") Then Stop: GoTo E
Dim J%
For J = 0 To jj.Siz_Ay(mAyRoot) - 1
    Dim mNmt$
    If jj.Fnd_ValFmSql(mNmt, "Select NmTbl from [$Tbl] where Tbl=" & mAyRoot(J)) Then Stop: GoTo E
    Debug.Print mAyRoot(J); ": "; mNmt
    Dim mAyId&(): If jj.Fnd_AyChd_ByRoot(mAyId, "$TblR", mAyRoot(J), "Tbl", "TblTo", True) Then Stop: GoTo E
    Dim mAnt$(): If jj.Fnd_AyVFmSql(mAnt, "Select NmTbl from [$Tbl] where Tbl in (" & jj.ToStr_AyLng(mAyId) & ")") Then Stop: GoTo E
    Debug.Print jj.ToStr_Ays(mAnt)
    Debug.Print jj.ToStr_AyLng(mAyId)
Next
Exit Function
E: Fnd_AyChd_ByRoot_Tst = True
End Function
#End If
Function Fnd_AyCnoDta(oAnFld$(), oAyCno() As Byte, pRge As Range) As Boolean
'Aim: Find the oAyCno & oAyFld which have a valid column name @ pRge(0,1)
Const cSub$ = "Fnd_AyCnoDta"
Dim mCnoLas As Byte: If jj.Fnd_CnoLas(mCnoLas, pRge(0, 1)) Then ss.A 1: GoTo E
Dim mN%
Dim iCno As Byte
For iCno = pRge.Column To mCnoLas - pRge.Column + 1
    Dim mV: mV = pRge(0, iCno).Value
    If VarType(mV) = vbString Then
        Dim I%
        For I = 0 To mN - 1
            If oAnFld(I) = mV Then MsgBox "Dup Col: " & mV: GoTo E
        Next
        ReDim Preserve oAyCno(mN)
        ReDim Preserve oAnFld(mN)
        oAyCno(mN) = iCno
        oAnFld(mN) = mV
        mN = mN + 1
    End If
Next
If jj.Siz_Ay(oAyCno) = 0 Then MsgBox "No valid Column name", , cSub: GoTo E
Exit Function
E: Fnd_AyCnoDta = True
End Function
'Function Fnd_V(oV$, pV) As Boolean
'If VarType(pV) = vbString Then oV = oV: Exit Function
'Fnd_V = True
'End Function
Function Fnd_Cmd(oCmd$, oRno&, pWs As Worksheet, pTar As Range) As Boolean
'Aim: Find the command of current pTar cell
Const cSub$ = "Fnd_Cmd"
If pTar.Count <> 1 Then GoTo E
oRno = pTar.Row: 'If oRno < jj.g.cRnoDta Then GoTo E
Dim mCno%: mCno = pTar.Column
If pTar.Interior.Color <> jj.g.cColrCmd Then GoTo E
Dim mV
If pWs.Cells(2, mCno).Interior.Color <> g.cColrCmd Then GoTo E
mV = pWs.Cells(2, mCno).Value
If VarType(mV) <> vbString Then GoTo E
If mV = "Cmd" Then mV = pTar.Value: If VarType(mV) <> vbString Then GoTo E
oCmd = Replace(Replace(Replace(mV, vbCr, ""), vbLf, ""), " ", "")
Exit Function
E: Fnd_Cmd = True
End Function
Function Fnd_RowAyVal(oAyVal$(), pAyRno&(), pWs As Worksheet, pRno&, pNm$) As Boolean
'Aim: Find {oAyVal} at {AyRno} of {pNm}.  Assume there is a name of x{pNm} defining each column
Const cSub = "RowAyVal"
On Error GoTo R
jj.Clr_Ays oAyVal
Dim mNm As Excel.Name
If jj.Fnd_Nm(mNm, pWs, "x" & pNm) Then ss.A 1: GoTo E
Dim mRge As Range: Set mRge = mNm.RefersToRange
Dim mCno As Byte: mCno = mRge.Column
Dim J%
For J = 0 To jj.Siz_Ay(pAyRno) - 1
    jj.Add_AyEle oAyVal, Nz(pWs.Cells(pAyRno(J), mCno).Value, "")
Next
GoTo X
R: ss.R
E: Fnd_RowAyVal = True: ss.B cSub, cMod, "AyRno,pWs,pRno,pNm", jj.ToStr_AyLng(pAyRno), jj.ToStr_Ws(pWs), pRno, pNm
X:
End Function
Function Fnd_Cno_ByNm(oCno As Byte, pWs As Worksheet, pNm$) As Boolean
Const cSub$ = "Fnd_Cno_ByNm"
Dim mNm As Excel.Name
If jj.Fnd_Nm(mNm, pWs, "x" & pNm) Then ss.A 1: GoTo E
oCno = mNm.RefersToRange.Column
Exit Function
R: ss.R
E: Fnd_Cno_ByNm = True: ss.B cSub, cMod, "pWs,pNm", jj.ToStr_Ws(pWs), pNm
End Function
Function Fnd_RowVal(oRowVal$, pWs As Worksheet, pRno&, pLn$) As Boolean
'Aim: Find {oRowVal} at {pRno} of list name in {pLn}.  Assume there is names of xXXX defining each column
Const cSub = "RowVal"
On Error GoTo R
oRowVal = ""
Dim mAn$(): mAn = Split(pLn, ",")
Dim J%
For J = 0 To jj.Siz_Ay(mAn) - 1
    Dim mCno As Byte: If jj.Fnd_Cno_ByNm(mCno, pWs, mAn(J)) Then ss.A 1: GoTo E
    Dim mV: mV = pWs.Cells(pRno, mCno).Value
    oRowVal = jj.Add_Str(oRowVal, CStr(mV))
Next
Exit Function
R: ss.R
E: Fnd_RowVal = True: ss.B cSub, cMod, "pWs,pRno,pLn", jj.ToStr_Ws(pWs), pRno, pLn
End Function
#If Tst Then
Function Fnd_RowVal_Tst() As Boolean
Dim mRowVal$: If jj.Fnd_RowVal(mRowVal, Worksheets("Tbl"), 10, "NmTbl,NmTy1xxx") Then Stop: GoTo E
Debug.Print mRowVal
Exit Function
E: Fnd_RowVal_Tst = True
End Function
#End If
Function Fnd_AyFfnRf(oAyFfnRf$(), pPrj As VBProject) As Boolean
'Aim: find {mAyFfnRf} of {pPrj}
Const cSub$ = "Fnd_AyFfnRf"
On Error GoTo R
ReDim oAyFfnRf(pPrj.References.Count - 1)
Dim iRf As VBIDE.Reference
Dim J%: J = 0
For Each iRf In pPrj.References
    oAyFfnRf(J) = iRf.FullPath: J = J + 1
Next
Exit Function
R: ss.R
E: Fnd_AyFfnRf = True: ss.B cSub, cMod, "pPrj", jj.ToStr_Prj(pPrj)
End Function
Function Fnd_Ffn_ByNmPrj(oFfnPrj$, pNmPrj$) As Boolean
Const cSub$ = "Fnd_Ffn_ByNmPrj"
Dim mPrj As VBProject: If jj.Fnd_Prj(mPrj, pNmPrj) Then ss.A 1: GoTo E
oFfnPrj$ = mPrj.Filename
Exit Function
E: Fnd_Ffn_ByNmPrj = True: ss.B cSub, cMod, "pNmPrj", pNmPrj
End Function
#If Tst Then
Function Fnd_Ffn_ByNmPrj_Tst() As Boolean
Dim mFfnPrj$: If jj.Fnd_Ffn_ByNmPrj(mFfnPrj, "jj") Then Stop: GoTo E
Debug.Print mFfnPrj
E: Fnd_Ffn_ByNmPrj_Tst = True
End Function
#End If
Function Fnd_An_BySetNm_Sql(oAn$(), pSetNm$, pSql$) As Boolean
'Aim: Find {oAn} by setting all first field of {pSql} if it like pSetNm$
Const cSub$ = "Fnd_An_BySetNm_Sql"
jj.Clr_Ays oAn
Dim mAyLik$(): If jj.Brk_Ln2Ay(mAyLik, pSetNm) Then ss.A 1: GoTo E
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, pSql) Then ss.A 2: GoTo E
With mRs
    While Not .EOF
        Dim mV$: mV = .Fields(0).Value
        If jj.IsLikAyLik(mV, mAyLik) Then jj.Add_AyEle oAn, mV
        .MoveNext
    Wend
End With
GoTo X
R: ss.R
E: Fnd_An_BySetNm_Sql = True: ss.B cSub, cMod, "pSetNm,pSql", pSetNm, pSql
X: jj.Cls_Rs mRs
End Function
#If Tst Then
Function Fnd_An_BySetNm_Sql_Tst() As Boolean
Dim mAn$(), mSetNm$, mSql$
mSql = "Select Distinct NmTbl from [$Tbl]"
mSetNm = "Typ*"
If jj.Fnd_An_BySetNm_Sql(mAn, mSetNm, mSql) Then Stop
Debug.Print Join(mAn, vbLf)
End Function
#End If
Function Fnd_Ws(oWs As Worksheet, pWb As Workbook, pNmWs$, Optional pSilent As Boolean) As Boolean
Const cSub$ = "Fnd_Ws"
On Error GoTo R
Set oWs = pWb.Sheets(pNmWs)
Exit Function
R: ss.R
E: Fnd_Ws = True: If Not pSilent Then ss.B cSub, cMod, "pWb,pNmWs", jj.ToStr_Wb(pWb), pNmWs
End Function
Function Fnd_RnoLas(oRnoLas&, pRge As Range) As Boolean
'Aim: find first empty cell of a column {pCno} in {pWs} starting {pRnoFm}into {oRnoLas}
Const cSub$ = "Fnd_RnoLas"
On Error GoTo R
If IsEmpty(pRge(1, 1).Value) Then oRnoLas = pRge.Row - 1: Exit Function
Dim mRge As Range: Set mRge = pRge(1, 1)
oRnoLas = mRge.End(xlDown).Row
Exit Function
R: ss.R
E: Fnd_RnoLas = True: ss.B cSub, cMod, "pRge", jj.ToStr_Rge(pRge)
End Function
Function Fnd_Ant_BySetNmt(oAnt$(), pSetNmt$, Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: Find {oAnt} by {pSetNmt} in {pDb}
Const cSub$ = "Fnd_Ant_BySetNmt"
jj.Clr_Ays oAnt
Dim mAyLikNmt$(): mAyLikNmt = Split(pSetNmt$, cComma)
Dim J%
For J = 0 To jj.Siz_Ay(mAyLikNmt) - 1
    Dim mAnt$(): If jj.Fnd_Ant_ByLik(mAnt, Trim(mAyLikNmt(J)), pDb) Then ss.A 1: GoTo E
    If jj.Add_AyAtEnd(oAnt, mAnt) Then ss.A 1: GoTo E
Next
Exit Function
R: ss.R
E: Fnd_Ant_BySetNmt = True: ss.B cSub, cMod, "pSetNmt,pDb", pSetNmt, jj.ToStr_Db(pDb)
End Function
#If Tst Then
Function Fnd_Ant_BySetNmt_Tst() As Boolean
Const cSub$ = "Fnd_Ant_BySetNmt_Tst"
Dim mSetNmt$: mSetNmt = "mst*,tbl*"
Dim mAntq$(): If jj.Fnd_Ant_BySetNmt(mAntq, mSetNmt) Then Stop
jj.Shw_Dbg cSub, cMod, "mSetNmt,Result(mAntq)", mSetNmt, jj.ToStr_Ays(mAntq)
End Function
#End If
Function Fnd_Antq_BySetNmtq(oAntq$(), pSetNmtq$, Optional pDb As DAO.Database = Nothing, Optional pQ$ = "") As Boolean
'Aim: Find {oAntq} by {pSetNmtq} in {pDb}
Const cSub$ = "Fnd_Antq_BySetNmtq"
jj.Clr_Ays oAntq
Dim mAyLikNmtq$(): mAyLikNmtq = Split(pSetNmtq$, cComma)
Dim J%

For J = 0 To jj.Siz_Ay(mAyLikNmtq) - 1
    Dim mAntq$(): If jj.Fnd_Antq_ByLik(mAntq, Trim(mAyLikNmtq(J)), pDb, pQ) Then ss.A 1: GoTo E
    If jj.Add_AyAtEnd(oAntq, mAntq) Then ss.A 1: GoTo E
Next
Exit Function
R: ss.R
E: Fnd_Antq_BySetNmtq = True: ss.B cSub, cMod, "pSetNmtq,pDb,pQ", pSetNmtq, jj.ToStr_Db(pDb), pQ
End Function
#If Tst Then
Function Fnd_Antq_BySetNmtq_Tst() As Boolean
Const cSub$ = "Fnd_Antq_BySetNmtq_Tst"
Dim mSetNmtq$: mSetNmtq = "mst*,tbl*"
Dim mAntq$(): If jj.Fnd_Antq_BySetNmtq(mAntq, mSetNmtq) Then Stop
jj.Shw_Dbg cSub, cMod, "mSetNmtq,Result(mAntq)", mSetNmtq, jj.ToStr_Ays(mAntq)
End Function
#End If
Function Fnd_AnTxtSpec(oAnTxtSpec$(), Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Fnd_AnTxtSpec"
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
If jj.Fnd_LoAyV_FmSql_InDb(mDb, "Select SpecName from MSysIMEXSpecs", "SpecName", oAnTxtSpec) Then
    Dim mA$(): oAnTxtSpec = mA
    ss.A 1: GoTo E
End If
Exit Function
R: ss.R
E: Fnd_AnTxtSpec = True: ss.B cSub, cMod, "pDb", jj.ToStr_Db(pDb)
End Function
Function Fnd_AnTxtSpec_Tst() As Boolean
Dim mAnTxtSpec$(): If jj.Fnd_AnTxtSpec(mAnTxtSpec) Then Stop
Debug.Print jj.ToStr_Ays(mAnTxtSpec)
End Function
Function Fnd_TxtSpecId(oTxtSpecId&, pNmSpec$, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Fnd_TxtSpecId"
jj.Set_Silent
If jj.Fnd_ValFmSql(oTxtSpecId, "Select SpecId from MSysIMEXSpecs where SpecName='" & pNmSpec & cQSng, pDb) Then GoTo E
GoTo X
E: Fnd_TxtSpecId = True
X: jj.Set_Silent_Rst
End Function
Function Fnd_AyDQry(oAyDQry() As jj.d_Qry, pNmqLik$, Optional pInclQDpd As Boolean = False, Optional pAcs As Access.Application = Nothing) As Boolean
'Aim: Find the {AyDQry} of {pNmQs} in {pDb}
Const cSub$ = "Fnd_AyDQry"
Dim mAcs As Access.Application: Set mAcs = jj.Cv_Acs(pAcs)
Dim mDb As DAO.Database:        Set mDb = mAcs.CurrentDb

Dim mAnq$(): If jj.Fnd_Anq_ByLik(mAnq, pNmqLik, mDb) Then ss.A 1: GoTo E
Dim N%: N = jj.Siz_Ay(mAnq)
If N = 0 Then
    Dim mAyDQry() As jj.d_Qry: oAyDQry = mAyDQry: Exit Function
    Exit Function
End If

Dim mLasMaj%: mLasMaj = -1
Dim J%, I%: I = 0
For J = 0 To N - 1
    Dim iQry As DAO.QueryDef: Set iQry = mDb.QueryDefs(mAnq(J))
        
    ReDim Preserve oAyDQry(I)
    Set oAyDQry(I) = New jj.d_Qry
    With oAyDQry(I)
        If .Brk_Nmqs(iQry.Name) Then ss.xx 1, cSub, cMod: GoTo Nxt
        .Typ = iQry.Type
        On Error Resume Next
        .Des = iQry.Properties("Description").Value
        On Error GoTo 0
        .Sql = iQry.Sql
        If mLasMaj <> .Maj Then
            If .Min <> 0 Then ss.A 1, "There is no Min Step 0", , "iQry.Name,pNmqLik,mLasMaj,mMaj", iQry.Name, pNmqLik, mLasMaj, .Maj: GoTo Nxt
            If .Typ <> DAO.QueryDefTypeEnum.dbQSelect Then ss.A 1, "The query of minor step 0 must be select query", , "The Query,Query Type(DAO.QueryDefTypeEnum)", iQry.Name, iQry.Type: GoTo Nxt
            mLasMaj = .Maj
        End If
    End With
    I = I + 1
Nxt:
Next
If pInclQDpd Then
    For J = 0 To N - 1
        With oAyDQry(J)
            .LnTbl = jj.ToStr_SqlLnt(.Sql)
        End With
    Next
End If
GoTo X
R: ss.R
E: Fnd_AyDQry = True: ss.B cSub, cMod, "pNmqLik,pInclQDpd,pAcs", pNmqLik, pInclQDpd, jj.ToStr_Acs(pAcs)
X:  Set mDb = Nothing
End Function
#If Tst Then
Function Fnd_AyDQry_Tst() As Boolean
Const cFfnCsv$ = "c:\aa.csv"
Const cDir$ = "D:\SPLHalfWayHouse\MPSDetail\VerNew@2007_01_04\WorkingDir\PgmObj\"

Dim mAyFb$(): If jj.Fnd_AyFn(mAyFb, cDir, "*.mdb") Then Stop
Dim I%

Dim mF As Byte: If jj.Opn_Fil_ForOutput(mF, cFfnCsv, True) Then Stop
Write #mF, "Mdb";
Dim mDQry As New d_Qry
If mDQry.WrtHdr(mF) Then Stop

For I = 0 To jj.Siz_Ay(mAyFb) - 1
    Dim mNmQs$: mNmQs = ""
    Select Case mAyFb(I)
    Case "MPS_GenDta.mdb":    mNmQs = "qryMPS"
    Case "MPS_GenRpt.mdb":    mNmQs = "qryMPS"
    Case "MPS_Odbc.mdb":      mNmQs = "qryOdbcMPS"
    Case "MPS_RfhCusGp.mdb":  mNmQs = "qryRfhCusGp"
    Case "RfhFc.mdb":     mNmQs = "qryFc,qryOdbcFc"
    End Select
    If Len(mNmQs) = 0 Then GoTo Nxt
    
    Dim mAnQs$(): mAnQs = Split(mNmQs, cComma)
    Dim mDb As DAO.Database:
    Do
        If jj.Opn_Db(mDb, cDir & mAyFb(I), True) Then Stop
        Dim N%: N = jj.Siz_Ay(mAnQs)
    
        Dim J%
        For J = 0 To N - 1
            Dim mAyDQry() As jj.d_Qry: If jj.Fnd_AyDQry(mAyDQry, mAnQs(J) & "*", True, mDb) Then Stop
            Dim K%
            For K = 0 To jj.Siz_AyDQry(mAyDQry) - 1
                If mAyDQry(K).Wrt(mF, mAyFb(I)) Then Stop
            Next
        Next
    Loop Until True
    mDb.Close
Nxt:
Next
Close #mF
Dim mWb As Workbook: If jj.Opn_Wb_RW(mWb, cFfnCsv) Then Stop
mWb.Application.Visible = True
End Function
#End If
Function Fnd_MaxLin%(pLines$)
Dim J%, L%, mAys$()
mAys = Split(pLines, vbLf)
For J = 0 To jj.Siz_Ay(mAys) - 1
    If L < Len(mAys(J)) Then L = Len(mAys(J))
Next
Fnd_MaxLin = L
End Function
Function Fnd_LnFld_ByNmtq(oLnFld$, pNmtq$, Optional pDb As DAO.Database = Nothing, Optional pInclTypFld As Boolean = False) As Boolean
'Aim: Find {oLnFld} by {pNmtq} in {pDb} to return if {pInclTypFld} & with {pSepChr}
Const cSub$ = "Fnd_LnFld_ByNmtq"
On Error GoTo R
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
If jj.IsTbl(pNmtq, mDb) Then oLnFld = jj.ToStr_Flds(mDb.TableDefs(pNmtq).Fields, pInclTypFld): Exit Function
If jj.IsQry(pNmtq, mDb) Then oLnFld = jj.ToStr_Flds(mDb.QueryDefs(pNmtq).Fields, pInclTypFld): Exit Function
ss.A 1, "Given pNmtq not exist in pDb"
GoTo E
R: ss.R
E: Fnd_LnFld_ByNmtq = True: ss.B cSub, cMod, "pNmtq,pDb", pNmtq, jj.ToStr_Db(pDb)
End Function
Function Fnd_LnFld_ByNmq(oLnFld$, pNmq$, Optional pInclTypFld As Boolean = False, Optional pSepChr$ = cComma, Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: Find {oLnFld} by {pNmq} in {pDb} to return if {pInclTypFld} & with {pSepChr}
'       Always look for any fields in {oLnFld} begins with yymd_.
'       If so, replace the field by:
'           Cdate(IIf(yymd_xxx=0,0,IIf(yymd_xxx=99999999,'9999/12/31',format(yymd_xxx,'0000\/00\/00')))) as xxx
'       Else
'           return oLnFld as "*"
Const cSub$ = "Fnd_LnFld_ByNmq"
On Error GoTo R
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
oLnFld = jj.ToStr_Flds(mDb.QueryDefs(pNmq).Fields, pInclTypFld, , pSepChr)
If Left(oLnFld, 3) = "Err" Then ss.A 1, "Cannot obtain field list from query", , "pNmq,Sql,oLnFld", pNmq, jj.ToSql_Nmq(pNmq), oLnFld: GoTo E
oLnFld = Cv_LnFld(oLnFld)
Exit Function
R: ss.R
E: Fnd_LnFld_ByNmq = True: ss.B cSub, cMod, "pNmq,pInclTypFld,pSepChr,pDb", pNmq, pInclTypFld, pSepChr, jj.ToStr_Db(pDb)
End Function
Function Fnd_LnFld_ByNmt(oLnFld$, pNmt$, Optional pInclTypFld As Boolean = False, Optional pSepChr$ = cComma, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Fnd_LnFld_ByNmt"
On Error GoTo R
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
oLnFld = jj.Cv_LnFld(jj.ToStr_Flds(mDb.TableDefs(jj.Rmv_SqBkt(pNmt)).Fields, pInclTypFld, , pSepChr))
Exit Function
R: ss.R
E: Fnd_LnFld_ByNmt = True: ss.B cSub, cMod, "pNmt,pInclTypFld,pSepChr", pNmt, pInclTypFld, pSepChr
End Function
#If Tst Then
Function Fnd_LnFld_ByNmt_Tst() As Boolean
Const cSub$ = "Fnd_LnFld_ByNmt_Tst"
Dim mLnFld$, mNmt$
Dim mResult As Boolean
Dim mCase As Byte: mCase = 1
Select Case mCase
Case 1: mNmt = "tmpFc_XlsKMR"
End Select
mResult = jj.Fnd_LnFld_ByNmt(mLnFld, mNmt)
jj.Shw_Dbg cSub, cMod, , "mLnFld, mNmt", mLnFld, mNmt
End Function
#End If
Function Fnd_Sffn_LgcMdb(oFbLgc$, pNmLgc$) As Boolean
Const cSub$ = "Fnd_Sffn_LgcMdb"
If Not jj.IsTbl("Av_LgcMdb") Then
    If jj.Crt_Tbl_FmLnkNmt(jj.Sdir_PgmObj & "Av.mdb", "Av_LgcMdb") Then ss.A 1: GoTo E
End If
If jj.Fnd_ValFmSql(oFbLgc, "Select FbLgc from Av_LgcMdb where NmLgc='" & pNmLgc & cQSng) Then ss.A 2: GoTo E
If Left(oFbLgc$, 2) = ".\" Then oFbLgc = jj.Sdir_PgmObj & mID(oFbLgc$, 3)
Exit Function
R: ss.R
E: Fnd_Sffn_LgcMdb = True: ss.B cSub, cMod, "pNmLgc"
End Function
#If Tst Then
Function Fnd_Sffn_LgcMdb_Tst() As Boolean
Dim mFbLgc$: If jj.Fnd_Sffn_LgcMdb(mFbLgc, "AddTbl") Then Stop
Debug.Print mFbLgc
End Function
#End If
Function Fnd_Sffn_LgcMdbTmp(oFbOldQsTmp$, pNmLgc$) As Boolean
Const cSub$ = "Fnd_Sffn_LgcMdbTmp"
Dim mFbLgc$: If jj.Fnd_Sffn_LgcMdb(mFbLgc, pNmLgc$) Then ss.A 1: GoTo E
oFbOldQsTmp = jj.Sdir_TmpLgc & "tmp" & jj.Nam_FilNam(mFbLgc$)
Exit Function
R: ss.R
E: Fnd_Sffn_LgcMdbTmp = True: ss.B cSub, cMod, "pNmLgc"
End Function
#If Tst Then
Function Fnd_Sffn_LgcMdbTmp_Tst() As Boolean
Dim mFbOldQsTmp$: If jj.Fnd_Sffn_LgcMdbTmp(mFbOldQsTmp, "AddTbl") Then Stop
Debug.Print mFbOldQsTmp
End Function
#End If
Function Fnd_AnQs(oAnQs$(), Optional pLikQry$ = "qry*", Optional pDb As DAO.Database = Nothing) As Boolean
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim mNmQsLas$, mNmQsCur$, N%, iQry As DAO.QueryDef
N = 0
For Each iQry In mDb.QueryDefs
    If Left(iQry.Name, 1) = "~" Then GoTo Nxt
    If Not iQry.Name Like pLikQry Then GoTo Nxt
    mNmQsCur = jj.Fnd_NmQs(iQry.Name)
    If mNmQsCur = "" Then GoTo Nxt
    If mNmQsLas <> mNmQsCur Then
        ReDim Preserve oAnQs(N)
        oAnQs(N) = mNmQsCur
        N = N + 1
        mNmQsLas = mNmQsCur
    End If
Nxt:
Next
If N = 0 Then jj.Clr_Ays oAnQs
End Function
#If Tst Then
Function Fnd_AnQs_Tst() As Boolean
Dim mAnQs$()
If jj.Fnd_AnQs(mAnQs) Then Stop
Debug.Print Join(mAnQs, vbLf)
End Function
#End If
Function Fnd_AnDpd(oAnDpd$(), pNmq$, Optional pAcs As Access.Application = Nothing) As Boolean
'Aim: The {AnDpd} by {pNmq}
Const cSub$ = "Fnd_AnDpd"
Dim mAcs As Access.Application: Set mAcs = jj.Cv_Acs(pAcs)
Dim mCase As Byte: mCase = 2
Select Case mCase
Case 1
    Dim mQry As Access.AccessObject: Set mQry = mAcs.CurrentData.AllQueries(pNmq)
    Dim mDpdInfo As Access.DependencyInfo: Set mDpdInfo = mQry.GetDependencyInfo
    Dim N%: N = mDpdInfo.Dependencies.Count
    If N = 0 Then jj.Clr_Ays oAnDpd: Exit Function
    Dim J%
    ReDim oAnDpd(N - 1)
    On Error Resume Next
    For J = 0 To N - 1
        oAnDpd(J) = "?"
        oAnDpd(J) = mDpdInfo.Dependencies(J).Name
    Next
Case 2
    Dim mSql$: mSql = mAcs.CurrentDb.QueryDefs(pNmq).Sql
    If jj.Fnd_SqlTbl(oAnDpd, mSql) Then ss.A 1: GoTo E
End Select
Exit Function
E: Fnd_AnDpd = True: ss.B cSub, cMod, "pNmq,pAcs", pNmq, jj.ToStr_Acs(pAcs)
End Function
#If Tst Then
Function Fnd_AnDpd_Tst() As Boolean
Dim mAnq$(), mAnt$()
If jj.Fnd_Anq_ByLik(mAnq, "*") Then Stop: GoTo E
Dim J%
For J = 0 To jj.Siz_Ay(mAnq) - 1
    Debug.Print mAnq(J),
    If jj.Fnd_AnDpd(mAnt, mAnq(J)) Then Stop: GoTo E
    Debug.Print jj.ToStr_Ays(mAnt)
Next
Exit Function
E: Fnd_AnDpd_Tst = True
End Function
#End If
Function Fnd_SqlTbl(oAnt$(), pSql$) As Boolean
ReDim oAnt(0): oAnt(0) = pSql
End Function
Function Fnd_Idx(oIdx%, pAy$(), pV$) As Boolean
'Aim: Find {pV} in {pAy} by return {oIdx}
Const cSub$ = "Fnd_Idx"
On Error GoTo R
For oIdx = 0 To jj.Siz_Ay(pAy) - 1
    If pAy(oIdx) = pV Then Exit Function
Next
oIdx = -1
Exit Function
R: ss.R
E: Fnd_Idx = True: ss.B cSub, cMod, cSub, cMod, "pAy,pV", jj.ToStr_Ays(pAy), pV
End Function
Function Fnd_IdxLng(oIdx%, pAyLng&(), pLng&) As Boolean
'Aim: Find {pLng} in {pAyLng} by return {oIdx}
Const cSub$ = "Fnd_IdxLng"
On Error GoTo R
For oIdx = 0 To jj.Siz_Ay(pAyLng) - 1
    If pAyLng(oIdx) = pLng Then Exit Function
Next
oIdx = -1
Exit Function
R: ss.R
E: Fnd_IdxLng = True: ss.B cSub, cMod, "pAyLng,pLng", jj.ToStr_AyLng(pAyLng), pLng
End Function
Function Fnd_IdxByt(oIdx%, pAyByt() As Byte, pByt As Byte) As Boolean
'Aim: Find {pLng} in {pAyLng} by return {oIdx}
Const cSub$ = "Fnd_IdxByt"
On Error GoTo R
For oIdx = 0 To jj.Siz_Ay(pAyByt) - 1
    If pAyByt(oIdx) = pByt Then Exit Function
Next
oIdx = -1
Exit Function
R: ss.R
E: Fnd_IdxByt = True: ss.B cSub, cMod, "pAyByt,pByt", jj.ToStr_AyByt(pAyByt), pByt
End Function
Function Fnd_FfnDtf$(pFbTar$, pNmtTar$)
Dim mDir$
If pFbTar = "" Then
    mDir = Fct.CurMdbDir & "DTF\"
Else
    mDir = Fct.Nam_DirNam(pFbTar) & "DTF\"
End If
jj.Crt_Dir mDir
Fnd_FfnDtf = mDir & pNmtTar & ".Dtf"
End Function
Function Fnd_Fn_By_Tp_n_CurFnn(oFn$, pTp$, pCurFnn$) As Boolean
Const cSub$ = "Fnd_Fn_By_Tp_n_CurFn"
'Aim: Assume pCurFnn is in fmt of xxxxx_nnnn  It is to return xxxxx_<tp>_Dta.mdb
Dim p%: p = InStrRev(pCurFnn, "_")
oFn = Left(pCurFnn, p) & pTp & "_Dta.mdb"
Exit Function
Fnd_Fn_By_Tp_n_CurFnn = True
End Function
#If Tst Then
Function Fnd_Fn_By_Tp_n_CurFnn_Tst() As Boolean
Dim mFn$: If jj.Fnd_Fn_By_Tp_n_CurFnn(mFn, "123", "xxxx_nnnn") Then Stop
Debug.Print mFn
End Function
#End If
Function Fnd_MaxEle(oIdx%, pAy$()) As Boolean
Dim N%: N = jj.Siz_Ay(pAy)
If N = 0 Then oIdx = -1: Exit Function
Dim J%, mMax$
oIdx = 0: mMax = pAy(0)
For J = 1 To N - 1
    If pAy(J) > mMax Then oIdx = J: mMax = pAy(J)
Next
End Function
Function Fnd_MinEle(oIdx%, pAy$()) As Boolean
Dim N%: N = jj.Siz_Ay(pAy)
If N = 0 Then oIdx = -1: Exit Function
Dim J%, mMin$
oIdx = 0: mMin = pAy(0)
For J = 1 To N - 1
    If pAy(J) < mMin Then oIdx = J: mMin = pAy(J)
Next
End Function
Function Fnd_LoQVal_ByFrm(oLoQVal$, pFrm As Access.Form, pAnCtl$()) As Boolean
Const cSub$ = "Fnd_LoQVal_ByFrm"
'Aim: Find {oLoQVal} by the control's NewValue in {pFrm} using {pAnCtl} as the control's name.
Dim J%, N%: N = jj.Siz_Ay(pAnCtl)
oLoQVal = ""
Dim mA$
For J = 0 To N - 1
    If jj.Fnd_QVal_ByFrm(mA, pFrm, pAnCtl(J)) Then ss.A 1: GoTo E
    oLoQVal = jj.Add_Str(oLoQVal, mA)
Next
Exit Function
R: ss.R
E: Fnd_LoQVal_ByFrm = True: ss.B cSub, cMod, "Frm,pAnCtl", jj.ToStr_Frm(pFrm), Join(pAnCtl, cComma)
End Function
#If Tst Then
Function Fnd_LoQVal_Frm_Tst() As Boolean
Const cNmFrm$ = "frmIIC_Tst"
Dim mFrm As Access.Form: If jj.Opn_Frm(cNmFrm, , , mFrm) Then Stop: GoTo E
Dim mAn$(): mAn = Split("ItemClass,Des,ICGL,ICALP5", cComma)
Dim mLstQVal$: If jj.Fnd_LoQVal_ByFrm(mLstQVal, mFrm, mAn) Then Stop: GoTo E
Debug.Print mLstQVal
Exit Function
E: Fnd_LoQVal_Frm_Tst = True
End Function
#End If
Function Fnd_LoAsg_InFrm(oLoAsg$, pFrm As Access.Form, pLm$, Optional oLoChgd$) As Boolean
Const cSub$ = "Fnd_LoAsg_InFrm"
'Aim: Find {oLoAsg} & {oLoChgd$} by the control's NewValue in {pFrm} using {pLm} as the control's name.
'     pLm     is fmt of aaa=xxx,bbb,ccc                            aaa,bbb,ccc will be used to list of control name. xxx,bbb,ccc will be used in {oLoAsg} & {oLoChgd}
'     oLoAsg  is fmt of xxx='nnnn',bbb=nnnn                     which is the [ssss] part of "Update tttt set ssss where wwww"
'     oLoChgd is fmt of xxx=[oooo]<--[nnnn]|bbb=[oooo]<--[nnnn] which will be show in status.
Dim mAnCtl$(), mAnAsg$(): If jj.Brk_Lm_To2Ay(mAnCtl, mAnAsg, pLm) Then ss.A 1: GoTo E
Dim J%, N%: N = jj.Siz_Ay(mAnCtl)
oLoAsg = "": oLoChgd = ""
Dim mA$, mB$
For J = 0 To N - 1
    If jj.Fnd_Asg_InFrm(mA, pFrm, mAnCtl(J), mAnAsg(J), , mB) Then ss.A 2: GoTo E
    If mA <> "" Then
        oLoAsg = jj.Add_Str(oLoAsg, mA)
        oLoChgd = jj.Add_Str(oLoChgd, mB, vbCrLf)
    End If
Next
Exit Function
R: ss.R
E: Fnd_LoAsg_InFrm = True: ss.B cSub, cMod, cSub, cMod, "Frm,pLm", jj.ToStr_Frm(pFrm), pLm$
End Function
#If Tst Then
Function Fnd_LoAsg_InFrm_Tst() As Boolean
Const cNmFrm$ = "frmIIC_Tst"
Dim mFrm As Access.Form: If jj.Opn_Frm(cNmFrm, , , mFrm) Then Stop: GoTo E
Dim mLm$: mLm = "ItemClass,Des=ICDES,ICGL,ICALP5"
Dim mLoAsg$, mLoChgd$: If jj.Fnd_LoAsg_InFrm(mLoAsg, mFrm, mLm, mLoChgd) Then Stop
Debug.Print jj.ToStr_NmV("mLoAsg", mLoAsg)
Debug.Print jj.ToStr_NmV("mLoChgd", mLoChgd)
Exit Function
E: Fnd_LoAsg_InFrm_Tst = True
End Function
#End If
Function Fnd_Asg_InFrm(oAsg$, pFrm As Access.Form, pNmCtl$, Optional pNmAsg$ = "", Optional pAlwNull As Boolean = False, Optional oChgd$) As Boolean
'Aim: return {oAsg},{oChgd} as ""                       if the Control of name {pNmCtl} in {pFrm} having equal .Value or .OldValue
'     else
'     return {oAsg} as mNmAsg=<QuotedValue>, and
'            {oChgd}as mNmAsg={OldVal}<--{NewVal} from the control {pNm} in {pFrm}.
'       Note:   If the new value is Null,
'                   if the field is num,str or bool, oAsg will return as mNmAsg=0 or '' or false
'                   other field type will return err.
Const cSub$ = "Fnd_Asg_InFrm"
On Error GoTo R
oAsg = "": oChgd = ""
Dim mNmAsg$: mNmAsg = NonBlank(pNmAsg, pNmCtl)
Dim mCtl As Access.Control: If jj.Fnd_Ctl(mCtl, pFrm, pNmCtl) Then ss.A 1: GoTo E
Dim mVNew, mVOld: mVNew = mCtl.Value: mVOld = mCtl.OldValue
If mVNew = mVOld Then Exit Function
Dim mTypSim As eTypSim: mTypSim = VarType(mVNew)
If mTypSim = vbNull Then
    If pAlwNull Then
        oAsg = mNmAsg & "=Null"
        oChgd = mNmAsg & "=Null<--[" & mVOld & "]"
        Exit Function
    End If
    Dim mRs As DAO.Recordset: Set mRs = pFrm.Recordset
    mTypSim = jj.Cv_TypDAO2Sim(mRs.Fields(pNmCtl).Type)
    Select Case mTypSim
    Case eTypSim_Bool: oAsg = mNmAsg & "=False": oChgd = mNmAsg & "=[False]<--[" & mVOld & "]"
    Case eTypSim_Num: oAsg = mNmAsg & "=0":      oChgd = mNmAsg & "=[0]<--[" & mVOld & "]"
    Case eTypSim_Str: oAsg = mNmAsg & "=''":     oChgd = mNmAsg & "=[]<--[" & mVOld & "]"
    Case Else
        ss.A 1, "The control having a null and it is not Bool,Num or Str", , "pFrm,pNmCtl,mNmAsg,SimTyp", jj.ToStr_Frm(pFrm), pNmCtl, mNmAsg, mTypSim
        GoTo E
    End Select
    Exit Function
End If
mTypSim = jj.Cv_V2Sim(mVNew)
Select Case mTypSim
Case eTypSim_Bool, eTypSim_Num, eTypSim_Str
    oAsg = mNmAsg & "=" & jj.Q_V(mVNew): oChgd = mNmAsg & "=[" & mVOld & "]<--[" & mVNew & "]"
Case Else
    ss.A 1, "The control having a value not being (Num,Bool,Str)", , "The Ctl's NewVal SimTyp", mTypSim
    GoTo E
End Select
Exit Function
R: ss.R
E: Fnd_Asg_InFrm = True: ss.B cSub, cMod, "pFrm,pNmCtl,mNmAsg", jj.ToStr_Frm(pFrm), pNmCtl, mNmAsg
End Function
Function Fnd_QVal_ByFrm(oQVal$, pFrm As Access.Form, pNmCtl$, Optional pAlwNull As Boolean = False) As Boolean
'Aim: Find {oQVal} as a quoted value for the control {pNmCtl}.Value in {pFrm}.
'     Only Num, Str or Bool type is allowed.
'     Null value will return 0, '' or False
Const cSub$ = "Fnd_QVal_ByFrm"
On Error GoTo R
Dim mV: mV = pFrm.Controls(pNmCtl).Value
Dim mTypSim As eTypSim

If VarType(mV) = vbNull Then
    If pAlwNull Then oQVal = "Null": Exit Function
    Dim mRs As DAO.Recordset: Set mRs = pFrm.Recordset
    mTypSim = jj.Cv_TypDAO2Sim(mRs.Fields(pNmCtl).Type)
    Select Case mTypSim
    Case eTypSim_Bool:  oQVal = "False"
    Case eTypSim_Num:   oQVal = "0"
    Case eTypSim_Str:   oQVal = "''"
    Case Else:          ss.A 1, "The control having a null and it is not Bool,Num or Str": GoTo E
    End Select
    Exit Function
End If
mTypSim = jj.Cv_V2Sim(mV)
Select Case mTypSim
Case eTypSim_Bool, eTypSim_Num, eTypSim_Str, eTypSim_Dte
    oQVal = Q_V(mV)
Case Else
    ss.A 2, "The control having a value not in (Num,Bool,Str,Dte)": GoTo E
End Select
Exit Function
R: ss.R
E: Fnd_QVal_ByFrm = True: ss.B cSub, cMod, "pFrm,pNmCtl,SimTyp of the ctl.value", jj.ToStr_Frm(pFrm), pNmCtl, mTypSim
End Function
Function Fnd_AyMacroStr_InStr(oAyMacroStr$(), pInStr$) As Boolean
Dim mP%, mA%, mB%, J%, mN As Byte
jj.Clr_Ays oAyMacroStr
mP = 1
Do
    mA = InStr(mP, pInStr, "{")
    If mA <= 0 Then Exit Function
    mB = InStr(mP + 1, pInStr, "}")
    If mB <= 0 Then Exit Function
    Dim mMacro$: mMacro = mID(pInStr, mA, mB - mA + 1)
    Dim mFnd As Boolean: mFnd = False
    For J = 0 To mN - 1
        If oAyMacroStr(J) = mMacro Then mFnd = True: Exit For
    Next
    If Not mFnd Then
        ReDim Preserve oAyMacroStr(mN)
        oAyMacroStr(mN) = mMacro: mN = mN + 1
    End If
    mP = mB + 1
Loop
End Function
#If Tst Then
Function Fnd_AyMacroStr_InStr_Tst() As Boolean
Dim mDtfTp$: If jj.Fnd_ResStr(mDtfTp, "DtfTp", True) Then Stop
Dim mAyMacroStr$(): If jj.Fnd_AyMacroStr_InStr(mAyMacroStr, mDtfTp) Then Stop
Debug.Print Join(mAyMacroStr, vbLf)
End Function
#End If
Function Fnd_Ffn_Fm_LnkXlsNmt(oFx$, pLnkXlsNmt$) As Boolean
'Aim: Find {oFx} from {pLnkXlsNmt} which is table name of a linked Excel
Const cSub$ = "Fnd_Ffn_Fm_LnkXlsNmt"
On Error GoTo R
Dim mCnn$: mCnn = CurrentDb.TableDefs(pLnkXlsNmt).Connect
If Left(mCnn, 10) <> "Excel 8.0;" Then ss.A 1, "Given pLnkXlsNmt does not have connection string starts with [Excel 8.0;]", , "pLnkXlsNmt,CnnStr", pLnkXlsNmt, mCnn: GoTo E
Dim mP%: mP = InStr(mCnn, "DATABASE="): If mP <= 0 Then ss.A 1, "Given pLnkXlsNmt connection string should be [DATABASE=]", , "pLnkXlsNmt,CnnStr", pLnkXlsNmt, mCnn: GoTo E
oFx = mID(mCnn, mP + 9)
Exit Function
R: ss.R
E: Fnd_Ffn_Fm_LnkXlsNmt = True: ss.B cSub, cMod, cSub, cMod
End Function
#If Tst Then
Function Fnd_Ffn_Fm_LnkXlsNmt_Tst() As Boolean
Const cSub$ = "Fnd_Ffn_Fm_LnkXlsNmt"
Const cFfn$ = "c:\Book1.xls"
Dim mWb As Workbook: If jj.Crt_Wb(mWb, cFfn, True) Then Stop
jj.Cls_Wb mWb, True
If jj.Crt_Tbl_FmLnkXls(cFfn) Then Stop
Dim mFfn$, mA$
mA = "Sheet1": If jj.Fnd_Ffn_Fm_LnkXlsNmt(mFfn, mA) Then Stop Else Debug.Print mA, mFfn
mA = "Sheet2": If jj.Fnd_Ffn_Fm_LnkXlsNmt(mFfn, mA) Then Stop Else Debug.Print mA, mFfn
mA = "Sheet3": If jj.Fnd_Ffn_Fm_LnkXlsNmt(mFfn, mA) Then Stop Else Debug.Print mA, mFfn
End Function
#End If
Function Fnd_Prm_FmTblPrm(oTrc&, oNmLgc$, Optional oLm$) As Boolean
'Aim: Find {oLn} & {oAyV} from tblPrm
'     Assume tblPrm has only 1 rec and is:Trc,NmLgc,Lm
Const cSub$ = "Fnd_Prm_FmTblPrm"
On Error GoTo R
With CurrentDb.OpenRecordset("Select * from tblPrm")
    oTrc = !Trc
    oNmLgc = !NmLgc
    oLm = Nz(!Lm, "")
    .Close
End With
Exit Function
R: ss.R
E: Fnd_Prm_FmTblPrm = True: ss.B cSub, cMod
End Function
Function Fnd_PrcDcl(oPrcDcl$, pMod$, pNmPrc$) As Boolean
'Aim: Get the 'Aim' lines into {oPrcDcl}.  Aim lines: first 50 lines with first line start with 'Aim and subsequent lines begin with '
Const cSub$ = "Fnd_PrcDcl"
On Error GoTo R
Const cMaxLen% = 250
Dim mS$
jj.Set_Silent
If jj.Fnd_PrcBody(mS, pMod, pNmPrc, , True) Then
    If jj.Fnd_PrcBody(mS, pNmPrc, pMod, , True) Then ss.A 1: GoTo E
End If
Dim mAy$(): mAy = Split(mS, vbCrLf)

'Put the Function Fnd_declaration lines into mAy() first
oPrcDcl = ""
Dim J%, I%, N%: N% = Fct.MinInt(50, jj.Siz_Ay(mAy) - 1)
For J = 0 To N
    oPrcDcl = jj.Add_Str(oPrcDcl, mAy(J), vbLf)
    If Right(mAy(J), 1) <> "_" Then Exit For
Next
'Find 'Aim
For J = J To N
    If Left(mAy(J), 4) = "'Aim" Then
        For I = J To N
            If Left(mAy(I), 1) <> cQSng Then GoTo X
            oPrcDcl = oPrcDcl & vbCrLf & mAy(I)
        Next
    End If
Next
GoTo X
R: ss.R
E: Fnd_PrcDcl = True: ss.C cSub, cMod, "pMod,pNmPrc", pMod, pNmPrc
X: jj.Set_Silent_Rst
End Function
Function Fnd_PrcDcl_Tst() As Boolean
Const cSub$ = "Fnd_PrcDcl_Tst"
Dim mPrcDcl$, mNmPrj_Nmm$, mNmPrc$
Dim mRslt As Boolean, mCase As Byte
mCase = 2
Select Case mCase
Case 1
    mNmPrj_Nmm = cLib & ".Fnd"
    mNmPrc = "PrcDcl"
Case 2
    mNmPrj_Nmm = cLib & ".Gen"
    mNmPrc = "Doc"
Case 3
    mNmPrj_Nmm = cLib & ".Read"
    mNmPrc = "Def_FmtTbl"
Case 4
    mNmPrj_Nmm = cLib & ".Bld"
    mNmPrc = "OdbcQs_ByAySelSql_ByDsn"
End Select
mRslt = jj.Fnd_PrcDcl(mPrcDcl, mNmPrj_Nmm, mNmPrc)
jj.Shw_DbgWin
Debug.Print mPrcDcl
End Function
Function Fnd_AyCno_XInRow(pWs As Worksheet, pRno&, pCnoFm As Byte, pCnoTo As Byte) As Byte()
Dim iCno As Byte, AyCno() As Byte, nCol As Byte
For iCno = pCnoFm To pCnoTo
    If pWs.Cells(pRno, iCno).Value = "X" Then
        nCol = nCol + 1
        ReDim Preserve AyCno(0 To nCol - 1)
        AyCno(nCol - 1) = iCno
    End If
Next
Fnd_AyCno_XInRow = AyCno()
End Function
Function Fnd_AyDir(oAyDir$(), pDir$) As Boolean
Const cSub$ = "Fnd_AyDir"
'History: Created on=2006/08/15; Modified on=2006/08/15
'Aim: Get a list of sub-dir in an Array (Start Index is 1) of a dir {pDir}
'==Start
If Not jj.IsDir(pDir) Then ss.A 1: GoTo E
Dim mSubDir$, AyLst$(), N As Byte
mSubDir = VBA.Dir(pDir & "*.*", vbDirectory)
While mSubDir <> ""
    If mSubDir <> "." And mSubDir <> ".." Then
        If GetAttr(pDir & mSubDir) And vbDirectory Then
            ReDim Preserve AyLst(0 To N)
            AyLst(N) = mSubDir
            N = N + 1
        End If
    End If
    mSubDir = VBA.Dir
Wend
oAyDir = AyLst
Exit Function
R: ss.R
E: Fnd_AyDir = True: ss.B cSub, cMod, "pDir", pDir
End Function
Function Fnd_AyFld(oAnFld$(), pNmtq$) As Boolean
Const cSub$ = "Fnd_AyFld"
Dim J As Byte
If jj.IsTbl(pNmtq) Then
    ReDim oAnFld(0 To CurrentDb.TableDefs(pNmtq).Fields.Count - 1)
    For J = 0 To CurrentDb.TableDefs(pNmtq).Fields.Count - 1
        oAnFld(J) = CurrentDb.TableDefs(pNmtq).Fields(J).Name
    Next
    Exit Function
End If
If jj.IsQry(pNmtq) Then
    ReDim oAnFld(0 To CurrentDb.QueryDefs(pNmtq).Fields.Count - 1)
    For J = 0 To CurrentDb.QueryDefs(pNmtq).Fields.Count - 1
        oAnFld(J) = CurrentDb.QueryDefs(pNmtq).Fields(J).Name
    Next
    Exit Function
End If
ss.A 1, "Given name is not table or query"
E: Fnd_AyFld = True: ss.B cSub, cMod, "pNmtq", pNmtq
End Function
Function Fnd_AyFn_ByLik(oAyFn$(), pDir$, pLik$, Optional pNoExt As Boolean = False) As Boolean
Const cSub$ = "Fnd_AyFn_ByLik"
'Aim: Fnd {oAyFn} by {pLik} in {pDir}
If Not jj.IsDir(pDir) Then ss.A 1: GoTo E
Dim mFn$, mAyFn$(), N As Byte
mFn = VBA.Dir(pDir & "*.*")
While mFn <> ""
    If mFn Like pLik Then
        ReDim Preserve mAyFn(N): N = N + 1
        If pNoExt Then
            mAyFn(N - 1) = jj.Cut_Ext(mFn)
        Else
            mAyFn(N - 1) = mFn
        End If
    End If
    mFn = VBA.Dir
Wend
oAyFn = mAyFn
Exit Function
R: ss.R
E: Fnd_AyFn_ByLik = True: ss.B cSub, cMod, ""
End Function
Function Fnd_AyFn(oAyFn$(), pDir$, Optional pFspc$ = "*.xls", Optional pNoExt As Boolean = False) As Boolean
'Aim: Fnd {oAyFn} by {pFSpc} in {pDir}
Const cSub$ = "Fnd_AyFn"
If Not jj.IsDir(pDir) Then ss.A 1: GoTo E
Dim mFn$, mAyFn$(), N As Byte, mAyLik$(): mAyLik = Split(pFspc, ",")
mFn = VBA.Dir(pDir & "*.*")
While mFn <> ""
    If jj.IsLikAyLik(mFn, mAyLik) Then
        ReDim Preserve mAyFn(N): N = N + 1
        If pNoExt Then
            mAyFn(N - 1) = jj.Cut_Ext(mFn)
        Else
            mAyFn(N - 1) = mFn
        End If
    End If
    mFn = VBA.Dir
Wend
oAyFn = mAyFn
Exit Function
R: ss.R
E: Fnd_AyFn = True: ss.B cSub, cMod, "pDir,pFspc,pNoExt"
End Function
Function Fnd_An2V_ByFrm(oAn2V() As tNm2V, pFrm As Access.Form, pLnFld$) As Boolean
'Aim: Fnd {oAn2V} from {pLnFld} in {pFrm} with optional to replace the {.Nm} of {oAn2V} by {pLnNew}
Const cSub$ = "Fnd_AnV_ByFrm"
Dim mAn_Frm$(), mAn_Host$(): If jj.Brk_Lm_To2Ay(mAn_Frm, mAn_Host, pLnFld) Then ss.A 1: GoTo E
Dim N%: N = jj.Siz_Ay(mAn_Frm): If N = 0 Then Exit Function
ReDim oAn2V(N - 1)
On Error GoTo R
Dim J%, iCtl As Access.Control
For J = 0 To N - 1
    If jj.Fnd_Ctl(iCtl, pFrm, mAn_Frm(J)) Then ss.A 2: GoTo E
    With oAn2V(J)
        .Nm = mAn_Host(J)
        .NewV = iCtl.Value
        .OldV = iCtl.OldValue
    End With
Next
Exit Function
R: ss.R
E: Fnd_An2V_ByFrm = True: ss.B cSub, cMod, "pFrm,pLnFld", jj.ToStr_Frm(pFrm), pLnFld
End Function
Function Fnd_An2V_ByFrm_Tst() As Boolean
Const cNmFrm$ = "frmSelBrandEnv"
If jj.Opn_Frm(cNmFrm) Then Stop: GoTo E
Dim mFrm As Access.Form: Set mFrm = Access.Application.Forms(cNmFrm)
Dim mAyNm2V() As tNm2V: If jj.Fnd_An2V_ByFrm(mAyNm2V, mFrm, "") Then Stop
Stop
Exit Function
E: Fnd_An2V_ByFrm_Tst = True
End Function
Function Fnd_Anm_ByPrj(oAnm$(), pPrj As VBProject _
    , Optional pLikNmm$ = "*" _
    , Optional pSrt As Boolean = False _
    ) As Boolean
Const cSub$ = "Fnd_Anm_ByPrj"
jj.Clr_Ays oAnm
With pPrj
    Dim mCmp As VBIDE.VBComponent
    For Each mCmp In .VBComponents
        If mCmp.Name Like pLikNmm Then jj.Add_AyEle oAnm, mCmp.Name
    Next
End With
If pSrt Then If jj.Srt_Ay(oAnm, oAnm) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: Fnd_Anm_ByPrj = True: ss.B cSub, cMod, "pPrj,pLikNmm,pSrt", jj.ToStr_Prj(pPrj), pLikNmm, pSrt
End Function
#If Tst Then
Function Fnd_Anm_Tst() As Boolean
Dim mAnPrj$(): If jj.Fnd_AnPrj(mAnPrj) Then Stop: GoTo E
Dim J%
For J = 0 To jj.Siz_Ay(mAnPrj) - 1
    Dim mPrj As VBProject: If jj.Fnd_Prj(mPrj, mAnPrj(J)) Then Stop: GoTo E
    Dim mAnm$(): If jj.Fnd_Anm_ByPrj(mAnm, mPrj) Then Stop: GoTo E
    Debug.Print mAnPrj(J) & ": " & jj.ToStr_Ays(mAnm)
Next
Exit Function
E: Fnd_Anm_Tst = True
End Function
#End If
Function Fnd_AnObj_ByPfx(oAnObj$(), pPfx$, Optional pTypObj As Access.AcObjectType = Access.AcObjectType.acQuery) As Boolean
Const cSub$ = "Fnd_AnObj_ByPfx"
If jj.Fnd_AnObj_ByPfx_InMdb(oAnObj, "", pPfx, pTypObj) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: Fnd_AnObj_ByPfx = True: ss.B cSub, cMod, "pPfx,pTypObj", pPfx, jj.ToStr_TypObj(pTypObj)
End Function
Function Fnd_AnObj_ByPfx_InMdb(oAnObj$(), pFb$, pPfx$, Optional pTypObj As Access.AcObjectType = Access.AcObjectType.acQuery) As Boolean
Const cSub$ = "Fnd_AnObj_ByPfx_InMdb"
Select Case pTypObj
Case Access.AcObjectType.acQuery
    If pFb = "" Or pFb = CurrentDb.Name Then
        If jj.Fnd_Anq_ByPfx(oAnObj, pPfx) Then ss.A 1: GoTo E
    Else
        Dim mDb As DAO.Database: If jj.Opn_Db_R(mDb, pFb) Then ss.A 2: GoTo E
        If jj.Fnd_Anq_ByPfx(oAnObj, pPfx, mDb) Then ss.A 3: GoTo E
        jj.Cls_Db mDb
    End If
    Exit Function
End Select
ss.A 4, "At this moment, only Query Type is supported": GoTo E
R: ss.R
E: Fnd_AnObj_ByPfx_InMdb = True: ss.B cSub, cMod, "pFb,pPfx,pTypObj", pFb, pPfx, jj.ToStr_TypObj(pTypObj)
End Function
Function Fnd_AnPrj(oAnPrj$() _
    , Optional pLikNmPrj$ = "*" _
    , Optional pSrt As Boolean = False _
    , Optional pAcs As Access.Application = Nothing _
    ) As Boolean
jj.Clr_Ays oAnPrj
Dim mAcs As Access.Application: Set mAcs = jj.Cv_Acs(pAcs)
Dim iPrj As VBProject
For Each iPrj In mAcs.VBE.VBProjects
    If iPrj.Name Like pLikNmPrj Then jj.Add_AyEle oAnPrj, iPrj.Name
Next
End Function
#If Tst Then
Function Fnd_AnPrj_Tst() As Boolean
Dim mAnPrj$(): If jj.Fnd_AnPrj(mAnPrj) Then Stop: GoTo E
Debug.Print jj.ToStr_Ays(mAnPrj, , vbLf)
jj.Shw_DbgWin
Exit Function
E: Fnd_AnPrj_Tst = True
End Function
#End If
Function Fnd_Anq_ByNmQs(oAnq$(), pNmQs$ _
    , Optional pMajBeg As Byte = 0 _
    , Optional pMajEnd As Byte = 99 _
    , Optional pDbQry As DAO.Database = Nothing) As Boolean
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDbQry)
Dim L As Byte: L = Len(pNmQs) + 1
Dim mNmQs$: mNmQs = pNmQs & "_"
jj.Clr_Ays oAnq

Dim mMajBeg$: mMajBeg$ = Format(pMajBeg, "00")
Dim mMajEnd$: mMajEnd$ = Format(pMajEnd, "00") & Chr(255)
Dim I%
Dim iQry As QueryDef: For Each iQry In mDb.QueryDefs
    If Left(iQry.Name, L) <> mNmQs Then GoTo Nxt
    If iQry.Name < pNmQs & "_" & mMajBeg$ Then GoTo Nxt
    If iQry.Name > pNmQs & "_" & mMajEnd$ Then Exit For
    ReDim Preserve oAnq(I): oAnq(I) = iQry.Name: I = I + 1
Nxt:
Next
End Function
Function Fnd_Anq_ByNmqs_Tst() As Boolean
Const cSub$ = "Fnd_Anq_ByNmqs_Tst"
Dim mAy$(), mNmQs$
Dim mResult As Boolean
Dim mCase As Byte: mCase = 1
Select Case mCase
Case 1
    mNmQs = "qryRfhInqAR"
    mResult = jj.Fnd_Anq_ByNmQs(mAy$, mNmQs, , 7)
End Select
jj.Shw_Dbg cSub, cMod, , "mResult,mNmqs,mAy", mResult, mNmQs, jj.ToStr_Ays(mAy)
End Function
Function Fnd_Anq_ByPfx(oAnq$(), pPfx$, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Fnd_Anq_ByPfx"
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim L%: L% = Len(pPfx)
Dim I%
jj.Clr_Ays oAnq
Dim mPfxXX$: mPfxXX = pPfx & Chr(255)
Dim iQry As DAO.QueryDef: For Each iQry In mDb.QueryDefs
    If Left(iQry.Name, L) = pPfx Then
        ReDim Preserve oAnq$(I)
        oAnq(I) = iQry.Name
        I = I + 1
    End If
    If iQry.Name > mPfxXX Then Exit For
Next
End Function
Function Fnd_Ant_ByLnk(oAnt_Lnk$(), Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Fnd_Ant_ByLnk"
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim I%
jj.Clr_Ays oAnt_Lnk
Dim iTbl As DAO.TableDef: For Each iTbl In mDb.TableDefs
    If iTbl.Connect <> "" Then
        ReDim Preserve oAnt_Lnk(I)
        oAnt_Lnk(I) = iTbl.Name
        I = I + 1
    End If
Next
End Function
#If Tst Then
Function Fnd_Ant_ByLnk_Tst() As Boolean
Dim mAnt_Lnk$()
If jj.Fnd_Ant_ByLnk(mAnt_Lnk) Then Stop
Debug.Print Join(mAnt_Lnk, vbLf)
End Function
#End If
Function Fnd_Ant_ByLik(oAnt$(), pLik$, Optional pDb As DAO.Database = Nothing, Optional pQ$ = "") As Boolean
Const cSub$ = "Fnd_Ant_ByLik"
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim I%
jj.Clr_Ays oAnt
Dim iTbl As DAO.TableDef: For Each iTbl In mDb.TableDefs
    Dim mNmt$: mNmt = iTbl.Name
    If Left(mNmt, 4) <> "MSYS" Then
        If iTbl.Name Like pLik$ Then
            ReDim Preserve oAnt$(I)
            oAnt(I) = jj.Q_S(iTbl.Name, pQ)
            I = I + 1
        End If
    End If
Next
End Function
Function Fnd_Anq_ByLik(oAnq$(), pLik$, Optional pDb As DAO.Database = Nothing, Optional pQ$ = "") As Boolean
Const cSub$ = "Fnd_Anq_ByLik"
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim I%
jj.Clr_Ays oAnq
Dim iQry As DAO.QueryDef: For Each iQry In mDb.QueryDefs
    Dim mNmq$: mNmq = iQry.Name
    If Left(mNmq, 1) <> "~" Then
        If mNmq Like pLik$ Then
            ReDim Preserve oAnq$(I)
            oAnq(I) = jj.Q_S(mNmq, pQ)
            I = I + 1
        End If
    End If
Next
End Function
Function Fnd_An_BySetNm(oAn$(), pAn$(), pSetNm$) As Boolean
'Aim:
End Function
Function Fnd_Antq_ByLik(oAntq$(), pLik$, Optional pDb As DAO.Database = Nothing, Optional pQ$ = "") As Boolean
Const cSub$ = "Fnd_Antq_ByLik"
Dim mDb As DAO.Database: Set pDb = jj.Cv_Db(mDb)
Dim I%
jj.Clr_Ays oAntq
Dim mAnt$(), mAnq$()
If jj.Fnd_Ant_ByLik(mAnt, pLik, mDb, pQ) Then ss.A 1: GoTo E
If jj.Fnd_Anq_ByLik(mAnq, pLik, mDb, pQ) Then ss.A 2: GoTo E
If jj.Add_Ay(oAntq, mAnt, mAnq) Then GoTo E
Exit Function
E: Fnd_Antq_ByLik = True: ss.B cSub, cMod, "pLik,pDb,pQ", pLik, jj.ToStr_Db(pDb), pQ
End Function
Function Fnd_AnWs_BySetWs(oAnWs$(), pWb As Workbook, Optional pSetWs$ = "*") As Boolean
Const cSub$ = "Fnd_AnWs_BySetWs"
On Error GoTo R
Dim mAnLikNmWs$(): mAnLikNmWs = Split(pSetWs, cComma)
Dim J%, mAnWs$()
For J = 0 To jj.Siz_Ay(mAnLikNmWs) - 1
    Dim mLikNmWs$: If jj.Fnd_AnWs_ByLikNmWs(mAnWs, pWb, Trim(mAnLikNmWs(J))) Then ss.A 1: GoTo E
    If jj.Add_AyAtEnd(oAnWs, mAnWs) Then ss.A 2: GoTo E
Next
Exit Function
R: ss.R
E: Fnd_AnWs_BySetWs = True: ss.B cSub, cMod, "pWb,pSetWs", jj.ToStr_Wb(pWb), pSetWs
End Function
Function Fnd_AnWs_ByLikNmWs(oAnWs$(), pWb As Workbook, pLikNmWs$) As Boolean
Const cSub$ = "Fnd_AnWs_ByLikNmWs"
On Error GoTo R
If InStr(pLikNmWs, "*") = 0 Then
    If jj.IsWs(pWb, pLikNmWs) Then
        ReDim oAnWs(0): oAnWs(0) = pLikNmWs
        Exit Function
    End If
    Dim mA$(): oAnWs = mA
    Exit Function
End If
Dim iWs As Worksheet, mN%: mN = 0
For Each iWs In pWb.Sheets
    If iWs.Name Like pLikNmWs Then
        ReDim Preserve oAnWs(mN)
        oAnWs(mN) = iWs.Name
        mN = mN + 1
    End If
Next
Exit Function
R: ss.R
E: Fnd_AnWs_ByLikNmWs = True: ss.B cSub, cMod, "pWb,pLikNmWs", jj.ToStr_Wb(pWb), pLikNmWs
End Function
Function Fnd_AnWs(oAnWs$(), pFx$, Optional pInclInvisible As Boolean = False) As Boolean
Const cSub$ = "Fnd_AnWs"
Dim mWb As Workbook, iWs As Worksheet, J As Byte
If jj.Opn_Wb(mWb, pFx, True) Then ss.A 1: GoTo E
If jj.Fnd_AnWs_ByWb(oAnWs, mWb, pInclInvisible) Then ss.A 2: GoTo E
mWb.Close False
Exit Function
R: ss.R
E: Fnd_AnWs = True: ss.B cSub, cMod, "pFx,pInclInvisible", pFx, pInclInvisible
End Function
Function Fnd_AnWs_ByWb(oAnWs$(), pWb As Workbook, Optional pInclInvisible As Boolean = False) As Boolean
Const cSub$ = "Fnd_AnWs"
On Error GoTo R
ReDim oAnWs$(pWb.Sheets.Count - 1)
Dim J%, iWs As Worksheet: J = 0
For Each iWs In pWb.Sheets
    If Not pInclInvisible And Not iWs.Visible Then GoTo Nxt
    oAnWs(J) = iWs.Name: J = J + 1
Nxt:
Next
Exit Function
R: ss.R
E: Fnd_AnWs_ByWb = True: ss.B cSub, cMod, "pWb,pInclinvisble", jj.ToStr_Wb(pWb), pInclInvisible
End Function
Public Function Fnd_AnWs_wColr(oAnWs$(), pWb As Workbook) As Boolean
'Aim: Find {oAnws} with color in tab
If TypeName(pWb) = "Nothing" Then Fnd_AnWs_wColr = True: Exit Function
Dim iWs As Worksheet, iCnt As Byte
For Each iWs In pWb.Sheets
    If iWs.Tab.Color Then
        ReDim Preserve oAnWs(iCnt)
        oAnWs(iCnt) = iWs.Name
        iCnt = iCnt + 1
    End If
Next
If iCnt = 0 Then Fnd_AnWs_wColr = True
End Function
Function Fnd_Brand_ById(oBrand$, pBrandId As Byte) As Boolean
Const cSub$ = "Fnd_Brand_ById"
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, "Select Brand from mstBrand where BrandId=" & pBrandId) Then ss.A 1: GoTo E
With mRs
    If .AbsolutePosition = -1 Then .Close: ss.A 2, "No record such {pBrandId} in mstBrand": GoTo E
    If Nz(!Brand, "") = "" Then .Close: ss.A 3, "Empty value in [Brand] field is found in mstBrand": GoTo E
    oBrand = !Brand
    .Close
End With
Exit Function
R: ss.R
E: Fnd_Brand_ById = True: ss.B cSub, cMod, "pBrandId", pBrandId
End Function
'Function Fnd_CdMd(oMod As CodeModule, pMod$) As Boolean
'Const cSub$ = "Fnd_CdMd"
'Dim mNmPrj$, mNmm$: If jj.Brk_Str_Both(mNmPrj, mNmm, pMod, ".") Then ss.A 1: GoTo E
'Dim mVBPrj As VBProject: If jj.Fnd_Prj(mVBPrj, mNmPrj) Then ss.A 2: GoTo E
'Dim mVBCmp As VBComponent: If jj.Fnd_VBCmp(mVBCmp, mVBPrj, mNmm) Then ss.A 3: GoTo E
'Set oMod = mVBCmp.CodeModule
'Exit Function
'R: ss.A 255,Err.Description, eException
'
'E: Fnd_
':ss.B cSub, cMod, ""
'    CdMd = True
'End Function
'Function Fnd_CdMod(oCdMod As CodeModule, pMod$) As Boolean
'Const cSub$ = "Fnd_CdMod"
'Dim mNmPrj$, mNmm$: If jj.Brk_Str_Both(mNmPrj, mNmm, pMod, ".") Then ss.A 1, "pMod must be xx.xx": GoTo E
'Dim mVBPrj As VBProject: If jj.Fnd_Prj(mVBPrj, mNmPrj) Then ss.A 2: GoTo E
'Dim iCmp As vbide.VBComponent: For Each iCmp In mVBPrj.VBComponents
'    If iCmp.Type = vbext_ct_StdModule Then If iCmp.Name = mNmm Then Set oCdMod = iCmp.CodeModule: Exit Function
'Next
'ss.A 3, "CdMod not found": GoTo E
'Exit Function
'R: ss.A 255,Err.Description, eException
'
'E: Fnd_
':ss.B cSub, cMod, "pMod", pMod
'    CdMod = True
'End Function
'Function Fnd_CdMod_Tst() As Boolean
'Const cSub$ = "Fnd_CdMod_Tst"
'Dim mCdMod As CodeModule, mMod$
'Dim mRslt As Boolean, mCase As Byte
'mCase = 1
'Select Case mCase
'Case 1
'    mMod$ = "jj.Fnd"
'End Select
'mRslt = jj.Fnd_CdMod(mCdMod, mMod$)
'jj.Shw_Dbg cSub, cMod, , "mRslt,mMod$", mRslt, mMod$
'End Function
Function Fnd_Cno_XInRow%(pWs As Worksheet, pRno&, Optional pLookFor$ = "X", Optional pCnoFm As Byte = 1, Optional pCnoTo As Byte = 255)
Dim iCno%, mStp%
mStp = IIf(pCnoTo >= pCnoFm, 1, -1)
For iCno = pCnoFm To pCnoTo Step mStp
    If pWs.Cells(pRno, iCno).Value = pLookFor$ Then Fnd_Cno_XInRow = iCno: Exit Function
Next
End Function
Function Fnd_Cno_EmptyCell_InRow(pWs As Worksheet _
    , Optional pRno& = 1 _
    , Optional pCnoFm% = 1 _
    , Optional pCnoTo% = 256 _
    ) As Byte
Dim iCno%, mStp%
mStp = IIf(pCnoTo >= pCnoFm, 1, -1)
For iCno = pCnoFm To pCnoTo Step mStp
    If IsEmpty(pWs.Cells(pRno, iCno).Value) Then Fnd_Cno_EmptyCell_InRow = iCno: Exit Function
Next
Fnd_Cno_EmptyCell_InRow = 0
End Function
Function Fnd_Ctl(oCtl As Access.Control, pFrm As Access.Form, pNmCtl$) As Boolean
Const cSub$ = "Fnd_Ctl"
On Error GoTo R
Set oCtl = pFrm.Controls(pNmCtl)
Exit Function
R: ss.R
E: Fnd_Ctl = True: ss.B cSub, cMod, "pFrm,pNmCtl", jj.ToStr_Frm(pFrm), pNmCtl
End Function
Function Fnd_Env_ById(oEnv$, pEnvId As Byte) As Boolean
Const cSub$ = "Fnd_Env_ById"
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, "Select Env from mstEnv where EnvId=" & pEnvId) Then ss.A 1: GoTo E
With mRs
    If .AbsolutePosition = -1 Then .Close: ss.A 2, "No such record {pEnvId} in mstEnv": GoTo E
    If Nz(!Env, "") = "" Then .Close: ss.A 3, "Empty value in [Env] field is found in mstEnv": GoTo E
    oEnv = !Env
    .Close
End With
Exit Function
R: ss.R
E: Fnd_Env_ById = True: ss.B cSub, cMod, "pEnvId", pEnvId
End Function
Function Fnd_FctNam_ByNmq$(pNmq$)
'Assume: Return QQQQ_n_xxxx_RunCode from {pNmq} of format QQQQ_nn_n_xxxx_RunCode
Dim mP1%: mP1 = InStr(pNmq, "_")
Dim mP2%: mP2 = InStr(mP1 + 1, pNmq, "_")
If mP1 = 0 Or mP2 = 0 Then Stop
If mP1 > mP2 Then Stop
If Right(pNmq, 8) = "_RUNCODE" Then
    Fnd_FctNam_ByNmq = Left(pNmq, mP1 - 1) & mID$(pNmq, mP2)
ElseIf Right(pNmq, 4) = "_Run" Then
    Dim mP3%: mP3 = InStr(mP2 + 1, pNmq, "_")
    If mP3 = 0 Then Stop
    Fnd_FctNam_ByNmq = Left(pNmq, mP1 - 1) & mID$(pNmq, mP3)
Else
    Stop
End If
End Function
#If Tst Then
Function Fnd_FctNam_ByNmq_Tst() As Boolean
Debug.Print Fnd_FctNam_ByNmq("qryABC_01_1_lsdf_RunCode")
End Function
#End If
Function Fnd_Ffn(oFfn$, Optional pDir$ = "C:\", Optional pFspc$ = "*.*", Optional pNmFSpc$ = "Any File", Optional pTit$ = "Select a file") As Boolean
Const cSub$ = "Fnd_Ffn"
With Application.FileDialog(msoFileDialogFilePicker)
    .InitialFileName = pDir
    .AllowMultiSelect = False
    .Title = pTit
    .Filters.Add pNmFSpc, pFspc
    .Show
    If .SelectedItems.Count = 1 Then oFfn = .SelectedItems(1): Exit Function
End With
E: Fnd_Ffn = True
End Function
#If Tst Then
Function Fnd_Ffn_Tst() As Boolean
Const cSub$ = "Fnd_Ffn_Tst"
Dim mFfn$: If jj.Fnd_Ffn(mFfn) Then Stop: GoTo E
jj.Shw_Dbg cSub, cMod, "mFfn", mFfn
Exit Function
E: Fnd_Ffn_Tst = True
End Function
#End If
Function Fnd_Fb_FmCnnStr(oFb$, pCnnStr$) As Boolean
Const cSub$ = "Fnd_Fb_FmCnnStr"
Const cDtaSrc$ = "Data Source="
'Provider=Microsoft.Jet.OLED4.0;User ID=Admin;Data Source=M:\07 ARCollection\ARCollection\WorkingDir\PgmObj\Template_ARInq.mdb;Mode=ReadWrite;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False
Dim mP1%: mP1 = InStr(pCnnStr, cDtaSrc)
Dim mP2%: mP2 = InStr(mP1, pCnnStr, ";")
If mP1 = 0 Or mP2 = 0 Or mP1 > mP2 Then ss.A 1, "Cannot find Data Source= or ; in given connection string": GoTo E
Dim L As Byte: L = Len(cDtaSrc)
oFb = mID(pCnnStr, mP1 + L, mP2 - mP1 - L)
Exit Function
E: Fnd_Fb_FmCnnStr = True: ss.B cSub, cMod, "pCnnStr", pCnnStr
End Function
Function Fnd_Fb_FmNmt_Lnk(oFb$, pNmt_Lnk$) As Boolean
Const cSub$ = "Fnd_Fb_FmNmt_Lnk"
Const cDb$ = "DATABASE="
'DATABASE=D:\SPLHalfWayHouse\MPSDetail\VerNew@2007_01_04\WorkingDir\PgmObj\MPS_RfhFc.Mdb;TABLE=tblFcPrm
On Error GoTo R
Dim L As Byte: L = Len(cDb)
Dim mCnnStr$: mCnnStr = CurrentDb.TableDefs(pNmt_Lnk).Connect
If Left(mCnnStr, 1) = ";" Then mCnnStr = mID(mCnnStr, 2)
If Left(mCnnStr, L) <> cDb Then ss.A 1, "pNmt_Lnk should have connect string started with " & cDb, , "pNmt_Lnk,CnnStr", pNmt_Lnk, mCnnStr: GoTo E
Dim mP%: mP = InStr(mCnnStr, ";")
If mP <= 0 Then oFb = mID(mCnnStr, L + 1): Exit Function
oFb = mID(mCnnStr, L + 1, mP - L)
Exit Function
R: ss.R
E: Fnd_Fb_FmNmt_Lnk = True: ss.B cSub, cMod, "pNmt_Lnk,mCnnStr", pNmt_Lnk, mCnnStr
End Function
#If Tst Then
Function Fnd_Fb_FmNmt_Lnk_Tst() As Boolean
Dim mFb$: If jj.Fnd_Fb_FmNmt_Lnk(mFb, "tblFcPrm") Then Stop
Debug.Print mFb
End Function
#End If
Function Fnd_FirstDateOfWk(pYr As Byte, pWk As Byte) As Date
Dim mFirstDateOfWk1 As Date
Select Case pYr
    Case 5:     mFirstDateOfWk1 = #1/2/2005#
    Case 6:     mFirstDateOfWk1 = #1/1/2006#
    Case 7:     mFirstDateOfWk1 = #1/7/2007#
    Case Else
        Stop
End Select
Fnd_FirstDateOfWk = mFirstDateOfWk1 + (pWk - 1) * 7
End Function
Function Fnd_FldVal_ByFld(oVal, pFld As DAO.Field) As Boolean
Const cSub$ = "Fnd_FldVal_ByFld"
On Error GoTo R
oVal = pFld.Value
Exit Function
R: ss.R
E: Fnd_FldVal_ByFld = True: ss.B cSub, cMod, "pFld", jj.ToStr_Fld(pFld)
End Function
Function Fnd_FldVal(oVal, pRs As DAO.Recordset, pNmFldRet$) As Boolean
Const cSub$ = "Fnd_FldVal"
On Error GoTo R
oVal = pRs.Fields(pNmFldRet).Value
Exit Function
R: ss.R
E: Fnd_FldVal = True: ss.B cSub, cMod, "pNmFldRet,pRs", pNmFldRet, jj.ToStr_Rs(pRs)
End Function
#If Tst Then
Function Fnd_FldVal_Tst() As Boolean
Dim mRs As DAO.Recordset: Set mRs = CurrentDb.OpenRecordset("Select * from mstBrand")
Dim mV: If jj.Fnd_FldVal(mV, mRs, "aaa") Then Stop
mRs.Close
End Function
#End If
Function Fnd_AyRgeRno(oAyRgeRno() As tRgeRno, pRge As Range) As Boolean
'Aim: find {oAyRgeRno} by each 'block'.  One 'block' one element in {oAyRgeRno}.
'     a 'block' is pCol for of RnoFm & RnoTo having same value.
'     pRge can be single cell, which means from this cell downward until empty cell
'     or a range of vertical cells to find the tRgeRno
Const cSub$ = "Fnd_AyRgeRno"
On Error GoTo R
'-- Find mNRow&
Dim mNRow&
If pRge.Count = 1 Then
    If pRge.End(xlDown).Row = 65536 Then
        ReDim oAyRgeRno(0)
        oAyRgeRno(0).Fm = pRge.Row
        oAyRgeRno(0).To = pRge.Row
        Exit Function
    End If
    mNRow = pRge.End(xlDown).Row - pRge.Row + 1
Else
    mNRow = pRge.SpecialCells(xlCellTypeLastCell).Row - pRge.Row + 1
End If

Dim iRno&, mRnoFm&: mRnoFm = pRge.Row
Dim mN%: mN = 0
Dim mWs As Worksheet: Set mWs = pRge.Parent
Dim mCno As Byte: mCno = pRge.Column
Dim mV, mVLas
mVLas = pRge.Cells(1, 1)
For iRno = pRge.Row To pRge.Row + mNRow - 1
    If mVLas <> mWs.Cells(iRno, mCno).Value Then
        mVLas = mWs.Cells(iRno, mCno).Value
        ReDim Preserve oAyRgeRno(mN)
        With oAyRgeRno(mN)
            .Fm = mRnoFm
            .To = iRno - 1
        End With
        mN = mN + 1
        mRnoFm = iRno
    End If
Next
If iRno > mRnoFm Then
    ReDim Preserve oAyRgeRno(mN)
    With oAyRgeRno(mN)
        .Fm = mRnoFm
        .To = iRno - 1
    End With
End If
Exit Function
R: ss.R
E: Fnd_AyRgeRno = True: ss.B cSub, cMod, "pRge", jj.ToStr_Rge(pRge)
End Function
Function Fnd_FmtDefSq(oFmtDefSq As tSq, pQt As QueryTable) As Boolean
Const cSub$ = "Fnd_FmtDefSq"
Const cTbl$ = "<Tbl>"

jj.Clr_Sq oFmtDefSq

oFmtDefSq.c1 = pQt.Destination.Column
Dim mWs As Worksheet: Set mWs = pQt.Parent

Dim mRgeRnoSearch As tRgeRno
With mRgeRnoSearch
    .To = pQt.Destination.Row - 1
    .Fm = .To - 30: If .Fm <= 0 Then .Fm = 1
    Dim iRno&: For iRno = .Fm To .To
        If mWs.Cells(iRno, oFmtDefSq.c1).Value = cTbl Then
            oFmtDefSq.r1 = iRno
            Dim jRno&: For jRno = iRno + 1 To .To
                If mWs.Cells(jRno, oFmtDefSq.c1).Value = cTbl Then
                    oFmtDefSq.r2 = jRno
                    Exit For
                End If
            Next
            Exit For
        End If
    Next
End With

With oFmtDefSq
    If .r1 = 0 Or .r2 = 0 Then ss.A 1, "No <Tbl> defintion": GoTo E
End With

Dim iCno As Byte: For iCno = oFmtDefSq.c1 + 1 To 255
    If mWs.Cells(oFmtDefSq.r1, iCno).Value = cTbl Then
        oFmtDefSq.c2 = iCno
        Exit For
    End If
Next
With oFmtDefSq
    If .c2 = 0 Then ss.A 2, "No <Tbl> defintion": GoTo E
End With
Exit Function
E: Fnd_FmtDefSq = True: ss.B cSub, cMod, "pQt", jj.ToStr_Qt(pQt)
End Function
Public Function Fnd_FreezedCell(oRno&, oCno As Byte, pWs As Worksheet) As Boolean
'Aim: Find the {oFreezedCellAdr} of {pWs}
Fnd_FreezedCell = True
pWs.Activate
Dim mWb As Workbook: Set mWb = pWs.Parent
Dim mWin As Window: Set mWin = pWs.Application.ActiveWindow
If mWin.Panes.Count <> 4 Then MsgBox "Given worksheet [" & pWs.Name & "] does not have 4 panes to find the Freezed Cell", vbCritical:  Exit Function
Dim mA$: mA = mWin.Panes(1).VisibleRange.Address
Dim mP%: mP = InStr(mA, ":")
mA = mID(mA, mP + 1)
Dim mRge As Range: Set mRge = pWs.Range(mA)
oRno = mRge.Row
oCno = mRge.Column
Fnd_FreezedCell = False
End Function
Private Function Fnd_FreezedCell_Tst() As Boolean
'Debug.Print Application.Workbooks.Count
'Debug.Print Application.Workbooks(1).FullName
'Dim mWs As Worksheet: Set mWs = Application.Workbooks(1).Sheets("Input - HKDP")
'Dim mRno&, mCno As Byte: If jj.Fnd_FreezedCell(mRno, mCno, mWs) Then Stop
'Debug.Print mRno & cComma & mCno
'Dim mSqLeft As cSq, mSqTop As cSq
End Function
Function Fnd_NxtBkFfnn(pFfnn$, oNxtBkFfnn$, oNxtBkNo As Byte) As Boolean
Dim mNmBk$: mNmBk = Right(pFfnn, 10)
If Left(mNmBk, 8) = " backup(" And Right(mNmBk, 1) = ")" Then
    oNxtBkNo = Val(mID(mNmBk, 9, 1)) + 1
    oNxtBkFfnn = Left(pFfnn, Len(pFfnn) - 2) & oNxtBkNo & ")"
    Exit Function
End If
oNxtBkNo = 1
oNxtBkFfnn = pFfnn & " backup(1)"
End Function
Function Fnd_Id_ByBrand(oBrandId As Byte, pBrand$) As Boolean
Const cSub$ = "Fnd_Id_ByBrand"
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, "Select BrandId from mstBrand where Brand='" & pBrand & cQSng) Then ss.A 1: GoTo E
With mRs
    If .AbsolutePosition = -1 Then .Close: ss.A 2, "No such record {pBrand} in mstBrand": GoTo E
    If Nz(!BrandId, 0) = 0 Then .Close: ss.A 3, "0 in [Brand] field is found in mstBrand": GoTo E
    oBrandId = !BrandId
    .Close
End With
Exit Function
R: ss.R
E: Fnd_Id_ByBrand = True: ss.B cSub, cMod, "pBrand", pBrand
End Function
Function Fnd_Id_ByEnv(oEnvId As Byte, pEnv$) As Boolean
Const cSub$ = "Fnd_Id_ByEnv"
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, "Select EnvId from mstEnv where Env='" & pEnv & cQSng) Then ss.A 1: GoTo E
With mRs
    If .AbsolutePosition = -1 Then .Close: ss.A 2, "No record {pEnv} in mstEnv": GoTo E
    If Nz(!EnvId, 0) = 0 Then .Close: ss.A 3, "0 in [Env] field is found in mstEnv": GoTo E
    oEnvId = !EnvId
    .Close
End With
Exit Function
R: ss.R
E: Fnd_Id_ByEnv = True: ss.B cSub, cMod, "pEnv", pEnv
End Function
Function Fnd_LasDteOfLasWW(pDte As Date) As Date
Dim mWeekday As Byte: mWeekday = Weekday(pDte, vbSunday) ' Sunday count as first day of a week & week day of Sunday is 1 & Saturday (last date of a week) is 7
Fnd_LasDteOfLasWW = pDte - mWeekday
End Function
Function Fnd_Layout(oLnFld$, pWs As Worksheet) As Boolean
Const cSub$ = "Fnd_Layout"
On Error GoTo R
Dim mV: mV = pWs.Cells(1, 1).Value
If IsEmpty(mV) Then ss.A 1, "A1 cell is empty", "pWs": GoTo E
oLnFld = mV
Dim J%: For J = 2 To 255
    mV = pWs.Cells(1, J).Value
    If IsEmpty(mV) Then Exit Function
    oLnFld = oLnFld & cComma & mV
Next
Exit Function
R: ss.R
E: Fnd_Layout = True: ss.B cSub, cMod, "pWs", jj.ToStr_Ws(pWs)
End Function
Function Fnd_Lbl(oLbl As Access.Label, pCtl As Access.Control) As Boolean
Dim mFrm As Access.Form: Set mFrm = pCtl.Parent
Dim mNm$
mNm = pCtl.Name & "_Lbl": If jj.Fnd_Lbl_ByNm(oLbl, mFrm, mNm) Then GoTo E
Exit Function
E: Fnd_Lbl = True
End Function
Function Fnd_Lbl_ByNm(oLbl As Access.Label, pFrm As Access.Form, pNm$) As Boolean
Const cSub$ = "Fnd_Lbl_ByNm"
On Error GoTo R
Set oLbl = pFrm.Controls(pNm)
Exit Function
R: Fnd_Lbl_ByNm = True
End Function
Function Fnd_Lv_FmDistSql(oLv$, pDistSql$, Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: Fnd {oLv} by joinning the first field value of {pDistSql} in {pDb}
Const cSub$ = "Fnd_Lv_FmDistSql"
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
With mDb.OpenRecordset(pDistSql)
    oLv = ""
    While Not .EOF
        oLv = jj.Add_Str(oLv, Nz(.Fields(0).Value, "#Null#"), cComma)
        .MoveNext
    Wend
    .Close
End With
Exit Function
R: ss.R
E: Fnd_Lv_FmDistSql = True: ss.B cSub, cMod, "pDistSql, mDb", pDistSql, jj.ToStr_Db(mDb)
End Function
Function Fnd_Lv_FmIdxTbl(oLv$, pNmIdxTbl$, pNmFldRet$, Optional pLn$ = "", Optional pAv = Nothing) As Boolean
'Aim: Fnd {oLv} of {pNmFldRet} in {pNmIdxTbl} with filter of list of field of {pLn} with list of value in {pAv}
Const cSub$ = "Fnd_Lv_FmIdxTbl"
Dim mWhere$: If jj.Bld_Where(mWhere, pLn, pAv) Then ss.A 1: GoTo E
Dim mSql$: mSql = jj.Fmt_Str("Select distinct {0} from {1}{2}", pNmFldRet, pNmIdxTbl, mWhere)
If jj.Fnd_Lv_FmDistSql(oLv, mSql) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: Fnd_Lv_FmIdxTbl = True: ss.B cSub, cMod, "pNmIdxTbl,pNmFldRet,pLn,pAv", pNmIdxTbl, pNmFldRet, pLn, jj.ToStr_Vayv(pAv)
End Function
Function Fnd_LoAyV_FmRs(pRs As DAO.Recordset, pLnFld$, oAyV0, Optional oAyV1, Optional oAyV2, Optional oAyV3, Optional oAyV4, Optional oAyV5) As Boolean
Const cSub$ = "Fnd_LoAyV_FmRs"
If VarType(oAyV0) And vbArray = 0 Then ss.A 1, "oAyV0 must be an array": GoTo E
If Not IsMissing(oAyV1) Then If VarType(oAyV1) And vbArray = 0 Then ss.A 2, "oAyV1 must be an array": GoTo E
If Not IsMissing(oAyV2) Then If VarType(oAyV2) And vbArray = 0 Then ss.A 3, "oAyV2 must be an array": GoTo E
If Not IsMissing(oAyV3) Then If VarType(oAyV3) And vbArray = 0 Then ss.A 4, "oAyV3 must be an array": GoTo E
If Not IsMissing(oAyV4) Then If VarType(oAyV4) And vbArray = 0 Then ss.A 5, "oAyV4 must be an array": GoTo E
If Not IsMissing(oAyV5) Then If VarType(oAyV5) And vbArray = 0 Then ss.A 6, "oAyV5 must be an array": GoTo E
Dim mAnFld$():  mAnFld = Split(pLnFld, cComma)
Dim mNFld%:     mNFld = jj.Siz_Ay(mAnFld): If mNFld <= 0 Or mNFld > 6 Then ss.A 7, "pLnFld is invalid (at most 6 elements)": GoTo E
Dim mNRec%:     If jj.Fnd_RecCnt_ByRs(mNRec, pRs) Then ss.A 1: GoTo E
If mNRec% = 0 Then
    Dim mAy()
    oAyV0 = mAy
    oAyV1 = mAy
    oAyV2 = mAy
    oAyV3 = mAy
    oAyV4 = mAy
    oAyV5 = mAy
    Exit Function
End If
If jj.Chk_Struct_Rs(pRs, pLnFld) Then ss.A 1: GoTo E
ReDim oAyV0(mNRec - 1), oAyV1(mNRec - 1), oAyV2(mNRec - 1), oAyV3(mNRec - 1), oAyV4(mNRec - 1), oAyV5(mNRec - 1)

With pRs
    .MoveFirst
    Dim iRec%: iRec = 0
    While Not .EOF
        Dim J%
        For J = 0 To mNFld - 1
            Select Case J
            Case 0: oAyV0(iRec) = .Fields(mAnFld(J)).Value
            Case 1: oAyV1(iRec) = .Fields(mAnFld(J)).Value
            Case 2: oAyV2(iRec) = .Fields(mAnFld(J)).Value
            Case 3: oAyV3(iRec) = .Fields(mAnFld(J)).Value
            Case 4: oAyV4(iRec) = .Fields(mAnFld(J)).Value
            Case 5: oAyV5(iRec) = .Fields(mAnFld(J)).Value
            End Select
        Next
        .MoveNext
        iRec = iRec + 1
    Wend
End With
Exit Function
R: ss.R
E: Fnd_LoAyV_FmRs = True: ss.B cSub, cMod, "pLnFld,mNRec,mNFld,Rs", pLnFld, mNRec, mNFld, jj.ToStr_Rs(pRs)
End Function
Function Fnd_LoAyV_FmRs_Tst() As Boolean
Dim mRs As DAO.Recordset: Set mRs = CurrentDb.OpenRecordset("Select * from tblOdbcSql")
Dim mLn$:                 mLn = "NmQs,NmDl" ' ,Sql,Sql_LclMd"
Dim AnQs(), AnDl(), AySql(), AySql_LclMd()
Fnd_LoAyV_FmRs_Tst = jj.Fnd_LoAyV_FmRs(mRs, mLn, AnQs, AnDl, AySql, AySql_LclMd)
jj.Shw_Dbg "GetLoAyV_FmRs_Tst", cMod, , "AnQs,AnDl", jj.ToStr_AyV(AnQs), jj.ToStr_AyV(AnDl)
End Function
Function Fnd_LoAyV_FmSql(pSql$, pLn$, oAyV0, Optional oAyV1 As Variant, Optional oAyV2, Optional oAyV3, Optional oAyV4, Optional oAyV5) As Boolean
Fnd_LoAyV_FmSql = Fnd_LoAyV_FmSql_InDb(CurrentDb, pSql, pLn, oAyV0, oAyV1, oAyV2, oAyV3, oAyV4, oAyV5)
End Function
Function Fnd_LoAyV_FmSql_InDb(pDb As DAO.Database, pSql$, pLn$, oAyV0, Optional oAyV1 As Variant, Optional oAyV2, Optional oAyV3, Optional oAyV4, Optional oAyV5) As Boolean
Const cSub$ = "Fnd_LoAyV_FmRs"
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim mRs As DAO.Recordset: If jj.Fnd_Rs_BySql(mRs, pSql, mDb) Then ss.A 1: GoTo E
If jj.Fnd_LoAyV_FmRs(mRs, pLn, oAyV0, oAyV1, oAyV2, oAyV3, oAyV4, oAyV5) Then ss.A 2: GoTo E
mRs.Close
Exit Function
R: ss.R
E: Fnd_LoAyV_FmSql_InDb = True: ss.B cSub, cMod, "pDb,pSql,pLn", jj.ToStr_Db(pDb), pSql, pLn
End Function
Function Fnd_LnFldVal(pRs As DAO.Recordset, pLnFld$ _
    , Optional oV0 _
    , Optional oV1 _
    , Optional oV2 _
    , Optional oV3 _
    , Optional oV4 _
    , Optional oV5 _
    , Optional oV6 _
    , Optional oV7 _
    , Optional oV8 _
    , Optional oV9 _
    ) As Boolean
Const cSub$ = "Fnd_LnFldVal"
Dim mAnFld$(): If jj.Brk_Ln2Ay(mAnFld, pLnFld) Then ss.A 1: GoTo E
Dim N%: N = jj.Siz_Ay(mAnFld): If N > 10 Then ss.A 1, "No more than 10 fields can be return": GoTo E
On Error GoTo R
Dim J%: For J = 0 To N - 1
    Select Case J
    Case 0: oV0 = pRs.Fields(mAnFld(J)).Value
    Case 1: oV1 = pRs.Fields(mAnFld(J)).Value
    Case 2: oV2 = pRs.Fields(mAnFld(J)).Value
    Case 3: oV3 = pRs.Fields(mAnFld(J)).Value
    Case 4: oV4 = pRs.Fields(mAnFld(J)).Value
    Case 5: oV5 = pRs.Fields(mAnFld(J)).Value
    Case 6: oV6 = pRs.Fields(mAnFld(J)).Value
    Case 7: oV7 = pRs.Fields(mAnFld(J)).Value
    Case 8: oV8 = pRs.Fields(mAnFld(J)).Value
    Case 9: oV9 = pRs.Fields(mAnFld(J)).Value
    End Select
Next
Exit Function
R: ss.R
E: Fnd_LnFldVal = True: ss.B cSub, cMod, "pRs,pLnFld,N,J", jj.ToStr_Rs(pRs), pLnFld, N, J
End Function
Function Fnd_MaxDir$(pDir$)
'Aim: within the all dir in {pDir}, return the dir with Max name
Dim mMaxDir$
Dim mDir$: mDir = VBA.Dir(pDir, vbDirectory)
While mDir <> ""
    If mDir > mMaxDir Then mMaxDir = mDir
    mDir = VBA.Dir
Wend
Fnd_MaxDir = mMaxDir
End Function
Function Fnd_MaxFfn$(pDir$, pFfnSpec$)
'Aim: within the all files of {pFfnSpec} in {pDir}, return the file with Max name
Dim mMaxFfn$
Dim mFfn$: mFfn = VBA.Dir(pDir & pFfnSpec)
While mFfn <> ""
    If mFfn > mMaxFfn Then mMaxFfn = mFfn
    mFfn = VBA.Dir
Wend
Fnd_MaxFfn = mMaxFfn
End Function
Function Fnd_MaxVal(oMaxVal, pNmt$, pNmFldMax$, Optional pLExpr$ = "", Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Fnd_MaxVal"
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim mWhere$: If pLExpr <> "" Then mWhere = " Where " & pLExpr
On Error GoTo R
With mDb.OpenRecordset(jj.Fmt_Str("select Max({0}) from {1}{2}", pNmFldMax, jj.Q_SqBkt(pNmt), mWhere))
    oMaxVal = Nz(.Fields(0).Value, 0)
    .Close
End With
Exit Function
R: ss.R
E: Fnd_MaxVal = True: ss.B cSub, cMod, "pNmt,pLExpr,pDb", pNmt, pLExpr, jj.ToStr_Db(pDb)
End Function
Public Function Fnd_Nm(oNm As Excel.Name, pWs As Worksheet, pNm$) As Boolean
If Not Fnd_Nm_InWs(oNm, pWs, pNm) Then Exit Function
Fnd_Nm = Fnd_Nm_InWb(oNm, pWs.Parent, pNm)
End Function
Public Function Fnd_Nm_InWb(oNm As Excel.Name, pWb As Workbook, pNm$) As Boolean
Const cSub$ = "Fnd_Nm_InWb"
On Error GoTo R
Set oNm = pWb.Names(pNm)
Exit Function
R: ss.R
E: Fnd_Nm_InWb = True: ss.B cSub, cMod, "pWb,pNm", jj.ToStr_Wb(pWb), pNm
End Function
Public Function Fnd_Nm_InWs(oNm As Excel.Name, pWs As Worksheet, pNm$) As Boolean
Const cSub$ = "Fnd_Nm_InWs"
On Error GoTo R
Set oNm = pWs.Names(pNm)
Exit Function
R: ss.R
E: Fnd_Nm_InWs = True: ss.B cSub, cMod, "pWs,pNm", jj.ToStr_Ws(pWs), pNm
End Function
Function Fnd_NmQs$(pNmq$)
'Aim: If the given {pNmq} in format of XXXX_NN_N_xxxx, return XXXX else return ""
' Postition of first 3 "_"
Dim mP1 As Byte: mP1 = InStr(pNmq, "_"):         If mP1 <= 1 Then Exit Function
Dim mP2 As Byte: mP2 = InStr(mP1 + 1, pNmq, "_"): If mP2 <= 0 Then Exit Function
If mP2 - mP1 <> 3 Then Exit Function
'
Dim mNN$: mNN = mID$(pNmq, mP1 + 1, 2)
Dim mA$:  mA = Left(mNN, 1)
If "0" > mA Or mA > "9" Then Exit Function
mA = mID$(mNN, 2, 1)
If "0" > mA Or mA > "9" Then Exit Function

Dim mN$: mN = mID$(pNmq, mP2 + 1, 1)
mA = mN
If "0" > mA Or mA > "9" Then Exit Function
Fnd_NmQs = Left(pNmq, mP1 - 1)
End Function
Function Fnd_Nmqs_Tst() As Boolean
Debug.Print jj.Fnd_NmQs("qryXCmp")
End Function
Function Fnd_Nmt_FmQt(oNmt$, pQt As QueryTable) As Boolean
Const cSub$ = "Fnd_Nmt_FmQt"
Select Case pQt.CommandType
Case XlCmdType.xlCmdTable: oNmt = pQt.CommandText
Case XlCmdType.xlCmdSql: If jj.Fnd_Nmt_FmSql(oNmt, pQt.CommandText) Then ss.A 1: GoTo E
Case Else
    ss.A 1, "Unexpected CmdTyp in given Qt": GoTo E
End Select
Exit Function
R: ss.R
E: Fnd_Nmt_FmQt = True: ss.B cSub, cMod, "pQt", jj.ToStr_Qt(pQt)
End Function
Function Fnd_Nmt_FmSql(oNmt$, pSql$) As Boolean
'Aim Find {oNmt} from {pSql} by looking up the token after "From"
Const cSub$ = "Fnd_Nmt_FmSql"
pSql = Replace(Replace(pSql, vbLf, " "), vbCr, " ")
Dim mAy$(): mAy = Split(pSql, " ")
Dim J%: For J = 0 To UBound(mAy)
    If mAy(J) = "From" Then
        Dim I As Byte: For I = 1 To UBound(mAy)
            If mAy(J + I) <> "" Then oNmt = Trim(mAy(J + I)): Exit Function
        Next
        ss.A 1, "No non-empty element in mAy()", , "mAy", jj.ToStr_Ays(mAy, "[]"): GoTo E
    End If
Next
ss.A 1, "No From in pSql": GoTo E
Exit Function
E: Fnd_Nmt_FmSql = True: ss.B cSub, cMod, "pSql", pSql
End Function
Function Fnd_Nmt_FmSql_Tst() As Boolean
Dim mNmt$: If jj.Fnd_Nmt_FmSql(mNmt, "lkdf lsdkj from    ksd  sdlk") Then Stop
Debug.Print mNmt
Stop
End Function
Function Fnd_PrcBody_ByMd(oStr$, pMd As CodeModule, pNmPrc$ _
    , Optional pBodyOnly As Boolean = False _
    ) As Boolean
Const cSub$ = "Fnd_PrcBody_ByMd"
Dim mAnPrc$(): If jj.Fnd_AnPrc_ByMd(mAnPrc, pMd, pNmPrc, , True, pBodyOnly) Then ss.A 1: GoTo E
If jj.Siz_Ay(mAnPrc) = 0 Then ss.A 2, "pNmPrc is not found": GoTo E
On Error GoTo R
Dim iNmPrc$, iPrcLinBeg$, iPrcLinEnd$, iPrcNLin$
If jj.Brk_Str_To3Seg(iNmPrc, iPrcLinBeg, iPrcLinEnd, mAnPrc(0)) Then ss.A 3: GoTo E
Dim mNLin&
Dim mLinBeg&: mLinBeg = iPrcLinBeg
Dim mLinEnd&: mLinEnd = iPrcLinEnd
mNLin = mLinEnd - mLinBeg + 1
oStr = pMd.Lines(mLinBeg, mNLin)
Exit Function
R: ss.R
E: Fnd_PrcBody_ByMd = True: ss.C cSub, cMod, "pMd,pNmPrc,pBodyOnly", jj.ToStr_Md(pMd), pNmPrc$, pBodyOnly
End Function
Function Fnd_PrcBody(oStr$, pMod$, pNmPrc$ _
    , Optional pAcs As Access.Application = Nothing _
    , Optional pBodyOnly As Boolean = False _
    ) As Boolean
Const cSub$ = "Fnd_PrcBody"
On Error GoTo R
Dim mNmPrj$, mNmm$: If jj.Brk_Str_Both(mNmPrj, mNmm, pMod, ".") Then ss.A 1, "pMod must have a '.'": GoTo E
Dim mPrj As VBProject: If jj.Fnd_Prj(mPrj, mNmPrj, pAcs) Then ss.A 2: GoTo E
Dim mMd As CodeModule: If jj.Fnd_Md(mMd, mPrj, mNmm) Then ss.A 3: GoTo E
If jj.Fnd_PrcBody_ByMd(oStr, mMd, pNmPrc, pBodyOnly) Then ss.A 4: GoTo E
Exit Function
R: ss.R
E: Fnd_PrcBody = True: ss.C cSub, cMod, "pMod,pNmPrc,pAcs,pBodyOnly", pMod, pNmPrc$, jj.ToStr_Acs(pAcs), pBodyOnly
End Function
'--------------------
#If Tst Then
Function Fnd_PrcBody_Tst() As Boolean
Const cSub$ = "Fnd_PrcBody_Tst"
Dim mPrcBody$, mNmPrj_Nmm$, mNmPrc$, mFb$
Dim mRslt As Boolean, mCase As Byte
mCase = 4
For mCase = 5 To 6
    Select Case mCase
    Case 1: mNmPrc = "zzGenDoc_FmtQry"
    Case 2
        mFb = "P:\WorkingDir\PgmObj\JMtcDb.mdb"
        mNmPrj_Nmm = "JMtcDb.RunGenTbl"
        mNmPrc = "qryGenTbl_Crt_TblKey_Run"
    Case 3
        mFb = ""
        mNmPrj_Nmm = "jj.Fnd"
        mNmPrc = "PrcBody_Tst"
    Case 4
        mFb = ""
        mNmPrj_Nmm = "jj.Fnd"
        mNmPrc = "Prp"
    Case 5
        mFb = ""
        mNmPrj_Nmm = "jj.Acpt"
        mNmPrc = "Dte_Tst"
    Case 6
        mFb = ""
        mNmPrj_Nmm = "jj.Acpt"
        mNmPrc = "PkVal"
    End Select
Next
Dim mAcs As Access.Application: If jj.Cv_Acs_FmFb(mAcs, mFb) Then Stop: GoTo E
mRslt = jj.Fnd_PrcBody(mPrcBody, mNmPrj_Nmm, mNmPrc, mAcs)
jj.Shw_Dbg cSub, cMod, "mRslt,mNmPrj_Nmm,mNmPrc", mRslt, mNmPrj_Nmm, mNmPrc
Debug.Print "------"
Debug.Print jj.Q_MrkUp(mPrcBody, "PrcBody")
GoTo E
E: Fnd_PrcBody_Tst = True
X: If mFb <> "" Then jj.Cls_CurDb mAcs
End Function
#End If
Function Fnd_Prp$(pNm$, pTypObj As AcObjectType, pNmPrp$)
Const cSub$ = "Fnd_Prp"
On Error GoTo R
Select Case pTypObj
Case AcObjectType.acTable _
     , AcObjectType.acReport _
     , AcObjectType.acForm _
     , AcObjectType.acMacro
            Fnd_Prp = CurrentDb.Containers("Tables").Documents(pNm).Properties(pNmPrp).Value
Case AcObjectType.acQuery: Fnd_Prp = CurrentDb.QueryDefs(pNm).Properties(pNmPrp).Value
Case Else:  ss.A 1, "Invalid pTypObj": GoTo E
End Select
Exit Function
R: ' ss.R
E: ' ss.B cSub, cMod, "pNm,pTypObj,pNmPrd", pNm, jj.ToStr_TypObj(pTypObj), pNmPrp
End Function
Function Fnd_RecCnt_ByNmtq(oRecCnt&, pNmtq$, Optional pLExpr$ = "", Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Fnd_RecCnt_ByNmt"
On Error GoTo R
Dim mSql$: mSql = jj.Fmt_Str("select count(*) from [{0}]{1}", jj.Rmv_SqBkt(pNmtq), jj.Cv_Where(pLExpr))
If jj.Fnd_ValFmSql(oRecCnt, mSql, pDb) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: Fnd_RecCnt_ByNmtq = True: ss.B cSub, cMod, "pNmtq,pLExpr,pDb", pNmtq, pLExpr, jj.ToStr_Db(pDb)
End Function
#If Tst Then
Function Fnd_RecCnt_ByNmtq_Tst() As Boolean
Dim aa&
Debug.Print jj.Fnd_RecCnt_ByNmtq(aa, "mstAllBrand")
Debug.Print aa
'
Const cFb$ = "c:\aa.mdb"
Dim mDb As DAO.Database: If jj.Crt_Db(mDb, cFb, True) Then Stop
If jj.Crt_Tbl_FmLoFld("aa", "aa Text 10", , , mDb) Then Stop
Call mDb.Execute("Insert into aa values('abc')")
Call mDb.Execute("Insert into aa values('abc')")
If jj.Fnd_RecCnt_ByNmtq(aa, "aa", , mDb) Then Stop
Debug.Print aa
jj.Shw_DbgWin
End Function
#End If

Function Fnd_RecCnt_ByRs(oNRec%, pRs As DAO.Recordset) As Boolean
Const cSub$ = "Fnd_RecCnt_ByRs"
If jj.IsNothing(pRs) Then Exit Function
oNRec = 0
On Error GoTo R
With pRs
    If .AbsolutePosition = -1 Then Exit Function
    .MoveFirst
    While Not .EOF
        oNRec = oNRec + 1
        .MoveNext
    Wend
    .MoveFirst
End With
Exit Function
R: ss.R
E: Fnd_RecCnt_ByRs = True: ss.B cSub, cMod, "pRs", jj.ToStr_Rs_NmFld(pRs)
End Function
Function Fnd_ResStr(oStr$, pNmRes$, Optional pNmPrc_Nmm$ = "jj.modResStr") As Boolean
Const cSub$ = "Fnd_ResStr"
If jj.Fnd_PrcBody(oStr, pNmPrc_Nmm, pNmRes, , True) Then ss.A 1: GoTo E
oStr = jj.Cut_LastLin(jj.Cut_FirstLin(jj.Rmv_FirstChr(oStr)))
Exit Function
E: Fnd_ResStr = True: ss.B cSub, cMod, "pNmPrc_Nmm,pNmRes", pNmPrc_Nmm, pNmRes
End Function
#If Tst Then
Function Fnd_ResStr_Tst() As Boolean
Const cSub$ = "Fnd_ResStr_Tst"
Dim mNmRes$, mStr$
Dim mRslt As Boolean, mCase As Byte: mCase = 1
Select Case mCase
Case 1
    mNmRes = "zzGenDoc_FmtQry"
Case 2
    mNmRes = "DtfTp"
Case 3
    mNmRes = "GenDoc_FmtMod"
End Select
mRslt = jj.Fnd_ResStr(mStr, mNmRes)
jj.Shw_Dbg cSub, cMod, "mRslt,mNmRes,mStr", mRslt, mNmRes, mStr
End Function
#End If
#If Tst Then
Function Fnd_ResStr1_Tst() As Boolean
Dim mStr$
If jj.Fnd_ResStr(mStr, "Fnd_ResStr1_Tst", cMod) Then Stop
Debug.Print mStr
End Function
#End If
Function Fnd_RgeCno_InRow(oRgeCno As tRgeCno, pWs As Worksheet, pRno&, Optional pCnoFm As Byte = 1, Optional pCnoTo As Byte = 255) As Boolean
'Aim: Looking for 'Beg' & 'End' in {pRno}
Const cSub$ = "Fnd_RgeCno_InRow"
With oRgeCno
    .Fm = 0
    .To = 0
    Dim iCno As Byte: For iCno = pCnoFm To pCnoTo
        If pWs.Cells(pRno, iCno).Value = "Beg" Then oRgeCno.Fm = iCno
        If pWs.Cells(pRno, iCno).Value = "End" Then oRgeCno.To = iCno: Exit Function
    Next
End With
ss.A 1, "Given row does not contain pair of Beg/End"
E: Fnd_RgeCno_InRow = True: ss.B cSub, cMod, "pRno,pCnoFm,pCnoTo", pRno, pCnoFm, pCnoTo
End Function
Function Fnd_Rs_ByFilter(oRs As DAO.Recordset, pNmt$, pLExpr$) As Boolean
Const cSub$ = "Fnd_Rs_ByFilter"
On Error GoTo R
Dim mSql$: mSql = jj.Fmt_Str("Select * from {0} where {1}", jj.Q_SqBkt(pNmt), pLExpr)
Set oRs = CurrentDb.OpenRecordset(mSql)
Exit Function
R: ss.R
E: Fnd_Rs_ByFilter = True: ss.B cSub, cMod, "pNmt,pLExpr", pNmt, pLExpr
End Function
Function Fnd_Rs_BySql(oRs As DAO.Recordset, pSql$, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Fnd_Rs_BySql"
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
On Error GoTo R
Set oRs = pDb.OpenRecordset(pSql)
Exit Function
R: ss.R
E: Fnd_Rs_BySql = True: ss.B cSub, cMod, "pSql,pDb", pSql, jj.ToStr_Db(pDb)
End Function
Function Fnd_SegFmCmd_2(oA1$, oA2$) As Boolean
Const cSub$ = "Fnd_SegFmCmd_2"
Dim mCmd$: mCmd = Command()
Dim mA$(): mA() = Split(mCmd, cComma)
If jj.Siz_Ay(mA) <> 2 Then ss.A 1, "/Cmd is expected as {Nmrptsht},{NmSess} format.", , "mCmd", mCmd: GoTo E
oA1 = mA(0)
oA2 = mA(1)
Exit Function
E: Fnd_SegFmCmd_2 = True: ss.B cSub, cMod
End Function
Function Fnd_SegFmCmd_3(oA1$, oA2$, oA3$) As Boolean
Const cSub$ = "Fnd_SegFmCmd_3"
Dim mCmd$: mCmd = Command()
Dim mA$(): mA() = Split(mCmd, cComma)
If jj.Siz_Ay(mA) <> 3 Then ss.A 1, "/Cmd is expected as {Nmrptsht},{NmSess}.{xxx} format.", , "mCmd", mCmd: GoTo E
oA1 = mA(0)
oA2 = mA(1)
oA3 = mA(2)
Exit Function
E: Fnd_SegFmCmd_3 = True: ss.B cSub, cMod
End Function
Function Fnd_Sql_ByNmq(oSql$, pNmq$) As Boolean
On Error GoTo R
oSql = CurrentDb.QueryDefs(pNmq).Sql
Exit Function
R: ss.R
E: Fnd_Sql_ByNmq = True
End Function
Function Fnd_Str_FmTxtFil(oS$, pFz$) As Boolean
Const cSub$ = "Fnd_Str_FmTxtFil"
On Error GoTo R
Dim mFno As Byte: If jj.Opn_Fil_ForInput(mFno, pFz) Then ss.A 1: GoTo E
oS = ""
While Not EOF(mFno)
    Dim mL$: Line Input #mFno, mL
    oS = jj.Add_Str(oS, mL, vbCrLf)
Wend
Close #mFno
Exit Function
R: ss.R
E: Fnd_Str_FmTxtFil = True: ss.B cSub, cMod, "pFz", pFz
X:
    Close #mFno
End Function
Function Fnd_Tbl(oTbl As DAO.TableDef, pNmt$) As Boolean
Const cSub$ = "Fnd_Tbl"
On Error GoTo R
Set oTbl = CurrentDb.TableDefs(jj.Rmv_SqBkt(pNmt))
Exit Function
R: ss.R
E: Fnd_Tbl = True: ss.B cSub, cMod, "pNmt", pNmt
End Function
Public Function Fnd_TwoSq(oSqLeft As cSq, oSqTop As cSq, pWs As Worksheet) As Boolean
'Aim: Find 2Sq: {oSqLeft} & {oSqTop} by {pWs}, {pRno} & {pCno}.  {pRno} & {pCno} are bottom right corner of pane of the freezed window.
Const cSub$ = "Fnd_TwoSq"
If TypeName(oSqLeft) = "Nothing" Then Set oSqLeft = New cSq
If TypeName(oSqTop) = "Nothing" Then Set oSqTop = New cSq
'Find the Freeze Cell
Dim mFreezeRno&, mFreezeCno As Byte: If jj.Fnd_FreezedCell(mFreezeRno, mFreezeCno, pWs) Then ss.A 1: GoTo E
'Detect mLasRno&, mLasCno
Dim mLasRno&, mLasCno As Byte
Dim mRge As Range: Set mRge = pWs.Cells.SpecialCells(xlCellTypeLastCell)
mLasRno = mRge.Row
mLasCno = mRge.Column
'Work From LasRno to pRno+1 to find first non-empty row so that oSqLeft is find
Dim iRno&, iCno%, mIsEmpty As Boolean
For iRno = mLasRno + 1 To mFreezeRno + 1 Step -1
    mIsEmpty = True
    For iCno = 1 To mFreezeCno
        If Not IsEmpty(pWs.Cells(iRno, iCno).Value) Then mIsEmpty = False: Exit For
    Next
    If mIsEmpty Then Exit For
Next
With oSqLeft
    .Cno1 = 1
    .Cno2 = mFreezeCno
    .Rno1 = mFreezeRno + 1
    .Rno2 = iRno - 1
End With
'Work From LasCno to pCno+1 to find first non-empty column so that oSqTop is find
For iCno = mLasCno + 1 To mFreezeCno + 1 Step -1
    mIsEmpty = True
    For iRno = 1 To mFreezeRno
        If Not IsEmpty(pWs.Cells(iRno, iCno).Value) Then mIsEmpty = False: Exit For
    Next
    If mIsEmpty Then Exit For
Next
With oSqTop
    .Cno1 = mFreezeCno + 1
    .Cno2 = iCno - 1
    .Rno1 = 1
    .Rno2 = mFreezeRno
End With
Exit Function
E: Fnd_TwoSq = True: ss.B cSub, cMod, "pWs", jj.ToStr_Ws(pWs)
End Function
Private Function Fnd_TwoSq_Tst() As Boolean
'Debug.Print Application.Workbooks.Count
'Debug.Print Application.Workbooks(1).FullName
'Dim mWs As Worksheet: Set mWs = Application.Workbooks(1).Sheets("Total")
'Dim mSqLeft As cSq, mSqTop As cSq
'If TwoSq(mSqLeft, mSqTop, mWs) Then Stop
'Debug.Print "mSqLeft=" & mSqLeft.ToStr
'Debug.Print "mSqTop=" & mSqTop.ToStr
End Function
Function Fnd_TypPrmRpt(oTypPrmRpt As tRpt, pNmRptSht$) As Boolean
Const cSub$ = "Fnd_TypPrmRpt"
On Error GoTo R
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, "Select * from tblRpt where Nmrptsht='" & pNmRptSht & cQSng) Then ss.A 1: GoTo E
With mRs
    If .AbsolutePosition = -1 Then ss.A 1, "Given report not define in tblRpt": GoTo E
    oTypPrmRpt.NmRpt = !NmRpt
    oTypPrmRpt.FmtStr_FnTo = Nz(!FmtStr_FnTo, "")
    oTypPrmRpt.QryPrm = Nz(!QryPrm, "")
    oTypPrmRpt.LnwsRmv = Nz(!LnwsRmv, "")
    oTypPrmRpt.HidePfLst_ThisNmSess = Nz(!HidePfLst_ThisNmSess, "")
    oTypPrmRpt.HidePfLst_ThisSess = Nz(!HidePfLst_ThisSess, "")
    oTypPrmRpt.HidePfLst_OtherSess = Nz(!HidePfLst_OtherSess, "")
    oTypPrmRpt.NmDta = Nz(!NmDta, "")
    oTypPrmRpt.EachSql = Nz(!EachSql, "")
    oTypPrmRpt.EachNmFld = Nz(!EachNmFld, "")
    oTypPrmRpt.EachLnwsRmv = Nz(!EachLnwsRmv, "")
    oTypPrmRpt.EachHidePfLst_ThisSess = Nz(!EachHidePfLst_ThisSess, "")
    oTypPrmRpt.EachHidePfLst_OtherSess = Nz(!EachHidePfLst_OtherSess, "")
End With
GoTo X
R: ss.R
E: Fnd_TypPrmRpt = True: ss.B cSub, cMod, "pNmRptSht", pNmRptSht
X: jj.Cls_Rs mRs
End Function
Function Fnd_LvFmRs_Of1Rec(oLv$, pRs As DAO.Recordset, pLnFld$, Optional pSepChr$ = cCommaSpc) As Boolean
Const cSub$ = "Fnd_LvFmRs_Of1Rec"
On Error GoTo R
Dim mAnFld$(): mAnFld = Split(pLnFld, cComma)
Dim N%: N = jj.Siz_Ay(mAnFld)
Dim mV
oLv = ""
Dim J%: For J = 0 To N - 1
    If jj.Fnd_FldVal_ByFld(mV, pRs.Fields(J)) Then ss.A 1, "One of the field cannot Get Fld Val", , "J", J: GoTo E
    oLv = jj.Add_Str(oLv, jj.Q_V(mV), pSepChr)
Next
Exit Function
R: ss.R
E: Fnd_LvFmRs_Of1Rec = True: ss.B cSub, cMod, "pRs,pLnFld,pSepChr", jj.ToStr_Rs_NmFld(pRs), pLnFld, pSepChr
End Function
Function Fnd_LvFmRs(oLv$, pRs As DAO.Recordset, Optional pNmFld$ = "", Optional pQ$ = "", Optional pSepChr$ = cComma) As Boolean
'Aim: FInd {oLv} from all record of first field  <pRs>.<pNmFld> of each record in {pRs} into {oLv}
Const cSub$ = "Fnd_LvFmRs"
oLv = ""
On Error GoTo R
Dim mAyV(): If jj.Fnd_AyVFmRs(mAyV, pRs, pNmFld) Then ss.A 1: GoTo E
oLv = jj.Join_AyV(mAyV, pQ, pSepChr)
Exit Function
R: ss.R
E: Fnd_LvFmRs = True: ss.B cSub, cMod, "pRs,pNmFld", jj.ToStr_Rs(pRs), pNmFld
End Function
#If Tst Then
Function Fnd_LvFmRs_Tst() As Boolean
Const cSub$ = "Fnd_LvFmRs"
Dim mNmt$, mNmFld$, mLv$, mCase As Byte

jj.Shw_Dbg cSub, cMod
For mCase = 1 To 1
    Select Case mCase
    Case 1: mNmt = "mstBrand": mNmFld = "BrandId"
    Case 2
    Case 3
    End Select
    Dim mRs As DAO.Recordset: Set mRs = CurrentDb.TableDefs(mNmt).OpenRecordset
    If jj.Fnd_LvFmRs(mLv, mRs, mNmFld) Then Stop
    mRs.Close
    Debug.Print jj.ToStr_LpAp(vbLf, "mCase,mNmt,mNmFld,mLv", mCase, mNmt, mNmFld, mLv)
    Debug.Print "----"
Next
End Function
#End If
Function Fnd_AyVFmRs(oAyV, pRs As DAO.Recordset, Optional pNmFld$ = "") As Boolean
'Aim: Find the first field in {pRs} for each record in pRs into {oAyV}
Const cSub$ = "Fnd_AyVFmRs"
On Error GoTo R
With pRs
    Dim N%: N = 0
    While Not .EOF
        ReDim Preserve oAyV(N)
        oAyV(N) = pRs.Fields(0).Value: N = N + 1
        .MoveNext
    Wend
End With
If N = 0 Then Exit Function
Exit Function
R: ss.R
E: Fnd_AyVFmRs = True: ss.B cSub, cMod, "pRs", jj.ToStr_Rs_NmFld(pRs)
End Function
Function Fnd_AyVFmSql(oAyV, pSql$, Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: Find the all value of first field in {pSql} into {oAys}
Const cSub$ = "Fnd_AyVFmSql"
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, pSql, pDb) Then ss.A 1: GoTo E
Fnd_AyVFmSql = jj.Fnd_AyVFmRs(oAyV, mRs)
GoTo X
E: Fnd_AyVFmSql = True: ss.B cSub, cMod, "pSql", pSql
X: jj.Cls_Rs mRs
End Function
#If Tst Then
Function Fnd_AyVFmSql_Tst() As Boolean
Dim mAyFy$(): If jj.Fnd_AyVFmSql(mAyFy(), "Select Fy From MstFy") Then Stop: GoTo E
Debug.Print jj.ToStr_Ays(mAyFy)
Exit Function
E: Fnd_AyVFmSql_Tst = True
End Function
#End If
Function Fnd_ValFmSql(oVal, pSql$ _
    , Optional pDb As DAO.Database = Nothing _
    ) As Boolean
'Aim: a value from a 'scalar' {pSql} in {pDb}
Const cSub$ = "Fnd_ValFmSql"
On Error GoTo R
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, pSql, pDb) Then ss.A 1: GoTo E
With mRs
    If .AbsolutePosition = -1 Then ss.A 1, "no record": GoTo E
    oVal = .Fields(0).Value
End With
GoTo X
R: ss.R
E: Fnd_ValFmSql = True: ss.B cSub, cMod, "pSql,pDb", pSql, jj.ToStr_Db(pDb)
X: jj.Cls_Rs mRs
End Function
#If Tst Then
Function Fnd_ValFmSql_Tst() As Boolean
Dim mSql$: mSql = "Select * from [#OldPgm]"
Dim mA$: If jj.Fnd_ValFmSql(mA, mSql) Then Stop: GoTo E
Debug.Print mA
Exit Function
E: Fnd_ValFmSql_Tst = True
End Function
#End If
Function Fnd_ValFmSql2(oV1, oV2, pSql$ _
    , Optional pDb As DAO.Database = Nothing _
    ) As Boolean
Const cSub$ = "Fnd_ValFmSql2"
'Aim: Find first 2 values from the first record of {pSql}
On Error GoTo R
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, pSql, pDb) Then ss.A 1: GoTo E
With mRs
    If .AbsolutePosition = -1 Then GoTo E
    oV1 = .Fields(0).Value
    oV2 = .Fields(1).Value
End With
GoTo X
R: ss.R
E: Fnd_ValFmSql2 = True: ss.B cSub, cMod, "pSql,pDb", pSql, jj.ToStr_Db(pDb)
X:
    jj.Cls_Rs mRs
End Function
Function Fnd_ValFmSql3(oV1, oV2, oV3, pSql$ _
    , Optional pDb As DAO.Database = Nothing _
    ) As Boolean
Const cSub$ = "Fnd_ValFmSql3"
'Aim: Find first 3 values from the first record of {pSql}
On Error GoTo R
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, pSql, pDb) Then ss.A 1: GoTo E
With mRs
    If .AbsolutePosition = -1 Then GoTo E
    oV1 = .Fields(0).Value
    oV2 = .Fields(1).Value
    oV3 = .Fields(2).Value
End With
GoTo X
R: ss.R
E: Fnd_ValFmSql3 = True: ss.B cSub, cMod, "pSql,pDb", pSql, jj.ToStr_Db(pDb)
X:
    jj.Cls_Rs mRs
End Function
Function Fnd_ValFmSql4(oV1, oV2, oV3, oV4, pSql$ _
    , Optional pDb As DAO.Database = Nothing _
    ) As Boolean
Const cSub$ = "Fnd_ValFmSql4"
'Aim: Find first 4 values from the first record of {pSql}
On Error GoTo R
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, pSql, pDb) Then ss.A 1: GoTo E
With mRs
    If .AbsolutePosition = -1 Then GoTo E
    oV1 = .Fields(0).Value
    oV2 = .Fields(1).Value
    oV3 = .Fields(2).Value
    oV4 = .Fields(3).Value
End With
GoTo X
R: ss.R
E: Fnd_ValFmSql4 = True: ss.B cSub, cMod, "pSql,pDb", pSql, jj.ToStr_Db(pDb)
X:
    jj.Cls_Rs mRs
End Function
Function Fnd_ValFmSql5(oV1, oV2, oV3, oV4, oV5, pSql$ _
    , Optional pDb As DAO.Database = Nothing _
    ) As Boolean
Const cSub$ = "Fnd_ValFmSql5"
'Aim: Find first 5 values from the first record of {pSql}
On Error GoTo R
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, pSql, pDb) Then ss.A 1: GoTo E
With mRs
    If .AbsolutePosition = -1 Then GoTo E
    oV1 = .Fields(0).Value
    oV2 = .Fields(1).Value
    oV3 = .Fields(2).Value
    oV4 = .Fields(3).Value
    oV5 = .Fields(4).Value
End With
GoTo X
R: ss.R
E: Fnd_ValFmSql5 = True: ss.B cSub, cMod, "pSql,pDb", pSql, jj.ToStr_Db(pDb)
X:
    jj.Cls_Rs mRs
End Function
Function Fnd_ValFmTbl_ByWhere(oVal, pNmt$, pNmFldRet$, pWhere$) As Boolean
Const cSub$ = "Fnd_ValFmTbl_ByWhere"
On Error GoTo R
Dim mSql$: mSql = jj.Fmt_Str("Select {0} from {1} where {2}", pNmFldRet, jj.Q_SqBkt(pNmt), pWhere)
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, mSql) Then ss.A 1: GoTo E
With mRs
    If .AbsolutePosition = -1 Then ss.A 2, "No record found by given mWhere in table": GoTo E
    oVal = .Fields(pNmFldRet)
End With
GoTo X
R: ss.R
E: Fnd_ValFmTbl_ByWhere = True: ss.B cSub, cMod, "pNmt,pNmFldRet,pWhere", pNmt, pNmFldRet, pWhere
X: jj.Cls_Rs mRs
End Function
Function Fnd_VbCmp(oVbCmp As VBComponent, pVBPrj As VBProject, pNmCmp$) As Boolean
Const cSub$ = "Fnd_VBCmp"
On Error GoTo R
Set oVbCmp = pVBPrj.VBComponents(pNmCmp)
Exit Function
R: Fnd_VbCmp = True
End Function
#If Tst Then
Function Fnd_VbCmp_Tst() As Boolean
Dim mPrj As VBProject: If jj.Fnd_Prj(mPrj, "jj") Then Stop: GoTo E
Dim mVbCmp As VBComponent: If jj.Fnd_VbCmp(mVbCmp, mPrj, "Form_frmWaitFor") Then Stop: GoTo E
Stop
Exit Function
E: Fnd_VbCmp_Tst = True
End Function
#End If
Function Fnd_VbCmp_FmWs(oVbCmp As VBIDE.VBComponent, pWs As Worksheet) As Boolean
Const cSub$ = "Fnd_VBCmp_FmWs"
On Error GoTo R
Dim mNmWs$: mNmWs = pWs.Name
For Each oVbCmp In pWs.Application.VBE.ActiveVBProject.VBComponents
    If oVbCmp.Type = vbext_ct_Document Then
        If oVbCmp.Properties("Name").Value = pWs.Name Then Exit Function
    End If
Next
ss.A 1, "VBCmp not find for ws"
GoTo E
R: ss.R
E: Fnd_VbCmp_FmWs = True: ss.C cSub, cMod, "VBCmp not find for ws", "pWs", jj.ToStr_Ws(pWs)
End Function
Function Fnd_Md_ByNm(oMd As CodeModule, pMod$ _
    , Optional pAcs As Access.Application = Nothing _
    ) As Boolean
Const cSub$ = "Fnd_Md_ByMogd"
Dim mNmPrj$, mNmm$: If jj.Brk_Str_Both(mNmPrj, mNmm, pMod, ".") Then ss.A 1: GoTo E
Dim mPrj As VBProject: If jj.Fnd_Prj(mPrj, mNmPrj, pAcs) Then ss.A 2: GoTo E
If jj.Fnd_Md(oMd, mPrj, mNmm) Then ss.A 3: GoTo E
Exit Function
E: Fnd_Md_ByNm = True: ss.B cSub, cMod, "pMod,pAcs", pMod, jj.ToStr_Acs(pAcs)
End Function
#If Tst Then
Function Fnd_Md_ByNm_Tst() As Boolean
Dim mMd As CodeModule: If jj.Fnd_Md_ByNm(mMd, "jj.cSq") Then Stop
Debug.Print jj.ToStr_Md(mMd)
Stop
End Function
#End If
Function Fnd_Md(oMd As CodeModule, pPrj As VBProject, pNmm$ _
    ) As Boolean
Const cSub$ = "Fnd_Md"
On Error GoTo R
Dim iCmp As VBComponent
Set iCmp = pPrj.VBComponents(pNmm)
Set oMd = iCmp.CodeModule
Exit Function
GoTo E
R: ss.R
E: Fnd_Md = True: ss.C cSub, cMod, "pPrj,pNmm", jj.ToStr_Prj(pPrj), pNmm
End Function
Function Fnd_PrcRgeRno(oRgeRno As tRgeRno, pMod$, pNmPrc$) As Boolean
'Aim: Find line range {oRgeRno} of {pNmPrc} in {pMd}
Const cSub$ = "Fnd_PrcRgeRno"
Dim mMd As CodeModule: If jj.Fnd_Md_ByNm(mMd, "jj.xFnd") Then ss.A 1: GoTo E
Fnd_PrcRgeRno = Fnd_PrcRgeRno_ByMd(oRgeRno, mMd, pNmPrc)
Exit Function
E: Fnd_PrcRgeRno = True: ss.A cSub, cMod, "pMod,pNmPrc", pMod, pNmPrc
End Function
#If Tst Then
Function Fnd_PrcRgeRno_Tst() As Boolean
jj.Shw_DbgWin
Dim mRgeRno As tRgeRno
If Fnd_PrcRgeRno(mRgeRno, "jj.xFnd", "Fnd_PrcRgeRno_Tst") Then Stop
Debug.Print mRgeRno.Fm, mRgeRno.To
If Fnd_PrcRgeRno(mRgeRno, "jj.xFnd", "Fnd_PrcRgeRno_ByMd") Then Stop
Debug.Print mRgeRno.Fm, mRgeRno.To
End Function
#End If
Function Fnd_PrcRgeRno_ByMd(oRgeRno As tRgeRno, pMd As CodeModule, pNmPrc$) As Boolean
'Aim: Find line range {oRgeRno} of {pMod}.{pNmPrc}
Const cSub$ = "Fnd_PrcRgeRno_ByMd"
On Error GoTo R
Dim mAnPrc_LinBeg_LinEnd$(): If Fnd_AnPrc_ByMd(mAnPrc_LinBeg_LinEnd, pMd, pNmPrc, , True) Then ss.A 1: GoTo E
Dim mN%: mN = jj.Siz_Ay(mAnPrc_LinBeg_LinEnd)
If mN = 0 Then oRgeRno.Fm = 0: oRgeRno.To = 0: Exit Function
If mN > 1 Then ss.A 2, "Return mAnPrc_LinBeg_LinEnd should be one element", , "mAnPrc_LinBeg_LinEnd", jj.ToStr_Ays(mAnPrc_LinBeg_LinEnd): GoTo E
Dim mNmPrc$, mLinBeg&, mLinEnd&: If jj.Brk_Str_To3Seg(mNmPrc, oRgeRno.Fm, oRgeRno.To, mAnPrc_LinBeg_LinEnd(0), ":") Then ss.A 2: GoTo E
Exit Function
R: ss.R
E: Fnd_PrcRgeRno_ByMd = True: ss.B cSub, cMod, "pMd,pNmPrc", jj.ToStr_Md(pMd), pNmPrc
End Function
Function Fnd_AnPrc_ByMd(oAnPrc_LinBeg_LinEnd$(), pMd As CodeModule _
    , Optional pLikNmPrc$ = "*" _
    , Optional pSrt As Boolean = False _
    , Optional pWithLinNo As Boolean = False _
    , Optional pBodyOnly As Boolean = False _
    ) As Boolean
'Aim: All lines begin and end being empty line or start with #.
Const cSub$ = "Fnd_AnPrc_ByMd"
On Error GoTo R
jj.Clr_Ays oAnPrc_LinBeg_LinEnd
With pMd
    Dim iLinNo&: iLinNo = .CountOfDeclarationLines + 1
    Dim iNmPrc$, iBeg&, iEnd&
    While iLinNo < .CountOfLines
        Dim mVBExt_pk_Proc As vbext_ProcKind
        iNmPrc = .ProcOfLine(iLinNo, mVBExt_pk_Proc)
        If iNmPrc Like pLikNmPrc Then
            
            If pWithLinNo Then
                iBeg = .ProcStartLine(iNmPrc, mVBExt_pk_Proc)
                iEnd = .ProcCountLines(iNmPrc, mVBExt_pk_Proc) + iBeg - 1
                Dim mBeg&, mEnd&, mL$, mA$
                mBeg = iBeg
                mEnd = iEnd
                For mBeg = iBeg To iEnd
                    mL = Trim(.Lines(mBeg, 1)): mA = Left(mL, 1)
                    If mL <> "" And mA <> cQSng And mA <> "#" Then Exit For
                Next
                
                For mEnd = iEnd To mBeg Step -1
                    mL = Trim(.Lines(mEnd, 1)): mA = Left(mL, 1)
                    If mL <> "" And mA <> cQSng And mA <> "#" Then Exit For
                Next
                If Not pBodyOnly Then
                    mL = Trim(.Lines(mBeg - 1, 1)): mA = Left(mL, 3)
                    If mA = "#If" Then mBeg = mBeg - 1
                    mL = Trim(.Lines(mEnd + 1, 1)): mA = Left(mL, 7)
                    If mA = "#End If" Then mEnd = mEnd + 1
                End If
                jj.Add_AyEle oAnPrc_LinBeg_LinEnd, iNmPrc & ":" & mBeg & ":" & mEnd
            Else
                jj.Add_AyEle oAnPrc_LinBeg_LinEnd, iNmPrc
            End If
        End If
        iLinNo = iLinNo + .ProcCountLines(iNmPrc, mVBExt_pk_Proc)
    Wend
End With
If pSrt Then If jj.Srt_Ay(oAnPrc_LinBeg_LinEnd, oAnPrc_LinBeg_LinEnd) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: Fnd_AnPrc_ByMd = True: ss.C cSub, cMod, "pMd,pLikNmPrc,pSrt", jj.ToStr_Md(pMd), pLikNmPrc, pSrt
End Function
#If Tst Then
Function Fnd_AnPrc_Tst() As Boolean
Const cSub$ = "Fnd_AnPrc_Tst"
Dim J%
Dim mPrj As VBProject:
Dim mAnm$(), mAnPrj$(), mAnPrc$(), mMd As CodeModule
Dim mCase As Byte
mCase = 2
Select Case mCase
Case 1
    If jj.Fnd_AnPrj(mAnPrj) Then Stop: GoTo E
    For J = 0 To jj.Siz_Ay(mAnPrj) - 1
        If jj.Fnd_Prj(mPrj, mAnPrj(J)) Then Stop: GoTo E
        If jj.Fnd_Anm_ByPrj(mAnm, mPrj) Then Stop: GoTo E
        Dim I%
        For I = 0 To jj.Siz_Ay(mAnm) - 1
            If jj.Fnd_Md(mMd, mPrj, mAnm(I)) Then Stop: GoTo E
            If jj.Fnd_AnPrc_ByMd(mAnPrc, mMd) Then Stop: GoTo E
            Debug.Print mAnPrj(J) & "." & mAnm(I) & ": " & jj.ToStr_Ays(mAnPrc)
        Next
    Next
Case 2
    Dim mLikPrc$:   mLikPrc = "qry*"
    Dim mNmPrj$:    mNmPrj = "JMtcDb"
    Dim mNmm$:      mNmm = "RunGentTbl"
    Dim mFbSrc$:    mFbSrc = "P:\WorkingDir\PgmObj\JMtcDb.mdb"
    Dim mAcs As Access.Application: If jj.Cv_Acs_FmFb(mAcs, mFbSrc) Then Stop: GoTo E
    If jj.Fnd_Prj(mPrj, mNmPrj, mAcs) Then Stop: GoTo E
    If jj.Fnd_Md(mMd, mPrj, mNmm) Then Stop: GoTo E
    If jj.Fnd_AnPrc_ByMd(mAnPrc, mMd, mLikPrc, , True) Then Stop
    Debug.Print jj.ToStr_Ays(mAnPrc, , vbLf)
End Select
Exit Function
R: ss.R
E: Fnd_AnPrc_Tst = True: ss.B cSub, cMod
End Function
#End If
Function Fnd_Prj(oPrj As VBProject, pNmPrj, Optional pApp As Application) As Boolean
Const cSub$ = "Fnd_Prj"
On Error GoTo R
Dim mApp As Application: Set mApp = jj.Cv_App(pApp)
Set oPrj = mApp.VBE.VBProjects(pNmPrj)
Exit Function
R: ss.R
E: Fnd_Prj = True: ss.C cSub, cMod, "pNmPrj,pAcs", pNmPrj, jj.ToStr_App(pApp)
End Function

