Attribute VB_Name = "xAdd"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xAdd"
Function Add_Fld(pTbl As DAO.TableDef, pFld As DAO.Field) As Boolean
Const cSub$ = "Add_Fld"
On Error GoTo R
pTbl.Fields.Append pFld
Exit Function
R: ss.R
E: Add_Fld = True: ss.B cSub, cMod, "pTbl,pFld", jj.ToStr_Tbl(pTbl), jj.ToStr_Fld(pFld)
End Function
Function Add_Am(oAm() As tMap, pAm1() As tMap, pAm2() As tMap) As Boolean
Const cSub$ = "Add_Am"
Dim mN1%: mN1 = jj.Siz_Am(pAm1)
Dim mN2%: mN2 = jj.Siz_Am(pAm2)
If mN1 = 0 And mN2 = 0 Then jj.Clr_Am oAm: Exit Function
ReDim oAm(mN1 + mN2 - 1)
Dim J%
For J = 0 To mN1 - 1
    oAm(J) = pAm1(J)
Next
For J = 0 To mN2 - 1
    oAm(J + mN1) = pAm2(J)
Next
On Error GoTo R
Exit Function
R: ss.R
E: Add_Am = True: ss.B cSub, cMod, "pAm1,pAm2", jj.ToStr_Am(pAm1), jj.ToStr_Am(pAm2)
End Function
#If Tst Then
Function Add_Am_Tst() As Boolean
Dim mAm1() As tMap: mAm1 = Get_Am_ByLm("A=1,B=2,C=3")
Dim mAm2() As tMap: mAm2 = Get_Am_ByLm("X=10,Y=20")
Dim mAm() As tMap: If jj.Add_Am(mAm, mAm1, mAm2) Then Stop: GoTo E
Debug.Print jj.ToStr_Am(mAm)
Shw_DbgWin
Exit Function
E:
End Function
#End If
Function Add_VbCmpToPrj(oVbCmp As VBIDE.VBComponent, pPrj As VBProject, pNmCmp$, Optional pTypCmp As VBIDE.vbext_ComponentType = VBIDE.vbext_ComponentType.vbext_ct_StdModule) As Boolean
Const cSub$ = "Add_VbCmpToPrj"
On Error GoTo R
If Not jj.Fnd_VbCmp(oVbCmp, pPrj, pNmCmp) Then Exit Function
Set oVbCmp = pPrj.VBComponents.Add(pTypCmp)
oVbCmp.Name = pNmCmp
Exit Function
R: ss.R
E: Add_VbCmpToPrj = True: ss.B cSub, cMod, "pPrj,pNmCmp,pTypCmp", jj.ToStr_Prj(pPrj), pNmCmp, jj.ToStr_TypCmp(pTypCmp)
End Function
#If Tst Then
Function Add_VbCmpToPrj_Tst() As Boolean
Dim mPrj As VBProject: If jj.Fnd_Prj(mPrj, "jj") Then Stop: GoTo E
Dim mVbCmp As VBComponent: If jj.Add_VbCmpToPrj(mVbCmp, mPrj, "xxx", vbext_ct_Document) Then Stop: GoTo E
Stop
Exit Function
E: Add_VbCmpToPrj_Tst = True
End Function
#End If
Function Add_MnuToPrj(pPrj As VBProject, pNmMnu$, Optional pCaption$ = "") As Boolean
'Aim: Add one userform {pNmMnu} to {pPrj} If menu exist, skip adding
Const cSub$ = "Add_MnuToPrj"
On Error GoTo R
Dim mVbCmp As VBComponent: If Add_VbCmpToPrj(mVbCmp, pPrj, pNmMnu, vbext_ct_MSForm) Then ss.A 1: GoTo E
If pCaption <> "" Then mVbCmp.Properties("Caption").Value = pCaption
mVbCmp.DesignerWindow.Close
Exit Function
R: ss.R
E: Add_MnuToPrj = True: ss.B cSub, cMod, "pPrj,pNmMnu", jj.ToStr_Prj(pPrj), pNmMnu
End Function
Function Add_MnuInWb(pWb As Workbook) As Boolean
'Aim: add one userform as WbMnu & each Ws add one userform as WsMnu.  If menu exist, skip adding
Const cSub$ = "Add_MnuInWb"
On Error GoTo R
Dim iWs As Worksheet
If Add_MnuToPrj(pWb.VBProject, "Mnu" & pWb.CodeName, pWb.CodeName & " Menu") Then ss.A 1: GoTo E
For Each iWs In pWb.Sheets
    If Add_MnuToPrj(pWb.VBProject, "Mnu" & iWs.CodeName, iWs.CodeName & " Menu") Then ss.A 2: GoTo E
Next

ReDim mAnWs$(pWb.Sheets.Count - 1)
Dim J%: J = 0
For Each iWs In pWb.Sheets
    mAnWs(J) = iWs.CodeName: J = J + 1
Next

Dim mNmPrj_Nmm$: mNmPrj_Nmm = "J" & mID(pWb.CodeName, 3) & ".g"

Dim mCode$

If jj.Fnd_ResStr(mCode, "Shw_MnuWb", True) Then ss.A 3: GoTo E
mCode = Trim(jj.Fmt_Str(mCode, pWb.CodeName))
If jj.Add_Prc(mNmPrj_Nmm, "Shw_MnuWb", mCode) Then ss.A 4: GoTo E

If jj.Fnd_ResStr(mCode, "Shw_MnuWs", True) Then ss.A 5: GoTo E
mCode = jj.Fmt_Str_Repeat_Ay_MultiLine(mCode, mAnWs)
If jj.Add_Prc(mNmPrj_Nmm, "Shw_MnuWs", mCode) Then ss.A 6: GoTo E

If jj.Fnd_ResStr(mCode, "ActWb_Set_MnuWb_MnuWs", True) Then ss.A 7: GoTo E
mNmPrj_Nmm = "J" & mID(pWb.CodeName, 3) & "." & pWb.CodeName
If jj.Add_Prc(mNmPrj_Nmm, "Workbook_WindowActivate", mCode) Then ss.A 6: GoTo E
Exit Function
R: ss.R
E: Add_MnuInWb = True: ss.B cSub, cMod, "pWb", jj.ToStr_Wb(pWb)
End Function
Function Add_MnuInXls() As Boolean
'Aim: Add mnu to each work book
Const cSub$ = "Add_MnuInXls"
On Error GoTo R
Dim iWb As Workbook
For Each iWb In Excel.Application.Workbooks
    If Add_MnuInWb(iWb) Then ss.A 1: GoTo E
Next
Exit Function
R: ss.R
E: Add_MnuInXls = True: ss.B cSub, cMod
End Function
Function Add_Prc(pMod$, pNmPrc$, pCode$, Optional pAcs As Access.Application = Nothing) As Boolean
'Aim: Add {pCode} to {pNmPrc} in {pMd}.  If {pNmPrc} exist, replace
Const cSub$ = "Add_Prc"
On Error GoTo R
Dim mMd As CodeModule: If jj.Fnd_Md_ByNm(mMd, pMod, pAcs) Then ss.A 1: GoTo E
Dim mLin&: If jj.Dlt_Prc_ByMd(mLin, mMd, pNmPrc) Then ss.A 2: GoTo E
mMd.InsertLines mLin, pCode: Exit Function
GoTo X
R: ss.R
E: Add_Prc = True
X:
End Function
#If Tst Then
Function Add_Prc_Tst() As Boolean
If jj.Add_Prc("jj.xAdd", "aaa", "aabbcc") Then Stop
End Function
#End If
Function Add_AyEleLng(ByRef oAyLng&(), pLng&, Optional pSilent As Boolean = False) As Boolean
'Aim: add pLng to oAyLng, return error if exist.
Const cSub$ = "Add_AyEleLng"
Dim N%, J%: N = jj.Siz_Ay(oAyLng)
For J = 0 To N - 1
    If oAyLng(J) = pLng Then ss.A 1, "pLng exist in oAyLng": GoTo E
Next
ReDim Preserve oAyLng(N): oAyLng(N) = pLng
Exit Function
E: Add_AyEleLng = True: If Not pSilent Then ss.B cSub, cMod, "oAyLng,pLng", jj.ToStr_AyLng(oAyLng), pLng
End Function
Function Add_Rf(pPrj As VBProject, pFfn$) As Boolean
'Aim: Add reference to {pPrj} from pFfn$, which is created by jj.Exp_Rf: .Name, .FullPath, .BuiltIn, .Type
'     The pPrj should be in other instance of pAcs, otherwise, adding Rf to it is not possible.
Const cSub$ = "Add_Rf"
On Error GoTo R
Dim mFno As Byte: If jj.Opn_Fil_ForInput(mFno, pFfn) Then ss.A 1: GoTo E
Dim mNmRf$, mFfnRf$, mIsBldIn As Boolean, mTypRf As VBIDE.vbext_RefKind
Dim mAyFfnRf$(): If jj.Fnd_AyFfnRf(mAyFfnRf, pPrj) Then ss.A 2: GoTo E
While Not EOF(mFno)
    Input #mFno, mNmRf, mFfnRf, mIsBldIn, mTypRf
    Dim mIdx%: If jj.Fnd_Idx(mIdx, mAyFfnRf, mFfnRf) Then ss.A 3: GoTo E
    If mIdx = -1 Then pPrj.References.AddFromFile mFfnRf
Wend
GoTo X
R: ss.R
E: Add_Rf = True: ss.B cSub, cMod, "pPrj,pFfn", jj.ToStr_Prj(pPrj), pFfn
X: Close #mFno
End Function
#If Tst Then
Function Add_Rf_Tst() As Boolean
Dim mCase As Byte
mCase = 2
Dim mFfn$, mPrj As VBProject:
Dim mAcs As Access.Application: Set mAcs = jj.g.gAcs
Dim mFb$: mFb = "c:\tmp\aa.mdb": If jj.Crt_Fb(mFb, True) Then Stop: GoTo E
If jj.Opn_CurDb(mAcs, mFb) Then Stop: GoTo E

Select Case mCase
Case 1
    Set mPrj = mAcs.VBE.ActiveVBProject
    Dim mRf As VBIDE.Reference: Set mRf = mPrj.References.AddFromFile("c:\program files\sap\frontend\sapgui\awkone.ocx")
    
    Dim mNmPrj$: mNmPrj = mPrj.Name
    mFfn = "C:\tmp\aa.reference.txt"
    If jj.Exp_Rf(mFfn, mNmPrj, mAcs) Then Stop: GoTo E

    mPrj.References.Remove mRf
Case 2
    mFfn = "P:\Documents\Pgm\jj.Reference.txt"
    Set mPrj = mAcs.VBE.ActiveVBProject
End Select
Stop
If jj.Add_Rf(mPrj, mFfn) Then Stop
mAcs.Visible = True
Stop
GoTo X
E: Add_Rf_Tst = True
X: jj.Cls_CurDb mAcs
   mAcs.Quit
   Set mAcs = Nothing
End Function
#End If
Function Add_Md_ToPrj(pPrj As VBProject, pFfn$) As Boolean
'Aim: create a module or class to {pPrj} from {pFfn}, which must be end with .bas or .cls
Const cSub$ = "Add_Md_ToPrj"
On Error GoTo R
If Not jj.IsFfn(pFfn) Then ss.A 1: GoTo E
Dim mFnn$, mDir$, mExt$: If jj.Brk_Ffn_To3Seg(mDir, mFnn, mExt, pFfn) Then ss.A 2: GoTo E
Dim mNmPrj$, mNmm$: If jj.Brk_Str_Both(mNmPrj, mNmm, mFnn, ".") Then ss.A 3: GoTo E
If pPrj.Name <> mNmPrj Then ss.A 4, "NmPrj not same as the pFfn": ss.A 4: GoTo E
Dim mTypCmp As VBIDE.vbext_ComponentType
Select Case mExt
Case ".bas": mTypCmp = vbext_ct_StdModule
Case ".cls": mTypCmp = vbext_ct_ClassModule
Case Else
    ss.A 2, "pFfn must be *.bas or *.cls": GoTo E
End Select

Dim mVbCmp As VBComponent: Set mVbCmp = pPrj.VBComponents.Add(mTypCmp)
With mVbCmp
    .Name = mNmm
    With .CodeModule
        .DeleteLines 1, .CountOfDeclarationLines
        .AddFromFile pFfn
    End With
End With
Exit Function
R: ss.R
E: Add_Md_ToPrj = True: ss.B cSub, cMod, "pPrj,pFfn", jj.ToStr_Prj(pPrj), pFfn
End Function
Function Add_Md(pAcs As Access.Application, pFfn$) As Boolean
'Aim: create a module or class in {pVBE} from {pFfn}, which must be end with .bas or .cls
Const cSub$ = "Add_Md"
If Not jj.IsFfn(pFfn) Then ss.A 1: GoTo E
Dim mFnn$, mDir$, mExt$: If jj.Brk_Ffn_To3Seg(mDir, mFnn, mExt, pFfn) Then ss.A 2: GoTo E
Dim mNmPrj$, mNmm$: If jj.Brk_Str_Both(mNmPrj, mNmm, mFnn, ".") Then ss.A 3: GoTo E
Dim mPrj As VBProject: Set mPrj = pAcs.VBE.ActiveVBProject
If mPrj.Name <> mNmPrj Then ss.A 4, "NmPrj not same as the pFfn": ss.A 4: GoTo E
If Left(mNmm, 5) = "Form_" Then
    Dim mFrm As Access.Form: Set mFrm = pAcs.CreateForm
    With mFrm.Module
        .DeleteLines 1, .CountOfDeclarationLines
        .AddFromFile pFfn
    End With
    Dim mNmFrm$: mNmFrm = mFrm.Name
    pAcs.DoCmd.Save acForm, mNmFrm
    pAcs.DoCmd.Close acForm, mNmFrm
    pAcs.DoCmd.Rename mID(mNmm, 6), acForm, mNmFrm
ElseIf Left(mNmm, 7) = "Report_" Then
    Dim mRpt As Access.Report: Set mRpt = pAcs.CreateReport
    With mRpt.Module
        .DeleteLines 1, .CountOfDeclarationLines
        .AddFromFile pFfn
    End With
    Dim mNmRpt$: mNmRpt = mRpt.Name
    pAcs.DoCmd.Save acReport, mNmRpt
    pAcs.DoCmd.Close acReport, mNmRpt
    pAcs.DoCmd.Rename mID(mNmm, 8), acReport, mNmRpt
Else
    If jj.Add_Md_ToPrj(mPrj, pFfn) Then ss.A 5: GoTo E
    pAcs.DoCmd.Save acModule, mNmm
End If
GoTo X
R: ss.R
E: Add_Md = True: ss.B cSub, cMod, "pAcs,pFfn", jj.ToStr_Acs(pAcs), pFfn
X:
End Function
Function Add_Tbl_ToTbl_ByNm(pNmtTo$, pNmtFm$, pNmFldTo$, Optional pNmFldFm$ = "") As Boolean
'Aim: Add Distinct pNmtFm!pNmFldFm into pNmtTo!pNmFldTo for those not exist
Const cSub$ = "Add_Tbl_ToTbl_ByNm"
Dim mNmFldFm$: mNmFldFm = jj.NonBlank(pNmFldFm, pNmFldTo)
Dim mSql$: mSql = jj.Fmt_Str("Insert into [{0}] ({1}) select Distinct {2} from [{3}]" & _
    " where {2} not in (Select {1} from [{0}])" & _
    " and {2}<>''", pNmtTo, pNmFldTo, mNmFldFm, pNmtFm)
If jj.Run_Sql(mSql) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: Add_Tbl_ToTbl_ByNm = True: ss.B cSub, cMod, "pNmtTo,pNmtFm,pNmFldTo,pNmFldFm", pNmtTo, pNmtFm, pNmFldTo, pNmFldFm
End Function
#If Tst Then
Function Add_Tbl_ToTbl_ByNm_Tst() As Boolean
If jj.Crt_Tbl_FmLoFld("#aa", "aa text 10") Then Stop: GoTo E
If jj.Crt_Tbl_FmLoFld("#bb", "bb text 10") Then Stop: GoTo E
If jj.Run_Sql("Insert into [#aa] values('1')") Then Stop: GoTo E
If jj.Run_Sql("Insert into [#aa] values('2')") Then Stop: GoTo E
If jj.Run_Sql("Insert into [#bb] values('2')") Then Stop: GoTo E
If jj.Run_Sql("Insert into [#bb] values('2')") Then Stop: GoTo E
If jj.Run_Sql("Insert into [#bb] values('3')") Then Stop: GoTo E
If jj.Run_Sql("Insert into [#bb] values('3')") Then Stop: GoTo E
If jj.Run_Sql("Insert into [#bb] values('3')") Then Stop: GoTo E
DoCmd.OpenTable "#aa"
Stop
If Add_Tbl_ToTbl_ByNm("#aa", "#bb", "aa", "bb") Then Stop
DoCmd.OpenTable "#aa"
Exit Function
E: Add_Tbl_ToTbl_ByNm_Tst = True
End Function
#End If
Function Add_AyEle(oAy$(), ByVal pEle$, Optional pKeepDup As Boolean = False) As Boolean
Dim N%: N = jj.Siz_Ay(oAy)
If Not pKeepDup Then
    Dim J%
    For J = 0 To N - 1
        If oAy(J) = pEle Then Exit Function
    Next
End If
ReDim Preserve oAy(N): oAy(N) = pEle: Exit Function
End Function
Function Add_AyByt(oAyByt() As Byte, ByVal pByt As Byte) As Boolean
Dim N%: N = jj.Siz_Ay(oAyByt)
ReDim Preserve oAyByt(N): oAyByt(N) = pByt: Exit Function
End Function
Function Add_Ws_ToWb(pWbTo As Workbook, pWbFm As Workbook, Optional pSetWs$ = "*") As Boolean
'Aim: Add {pWbFm}!{pSetWs} to {pWbTo}.  If ws exist, it will be replaced and position retended, else add to the end.
Const cSub$ = "Add_Ws_ToWb"
Dim mAnWs$(): If jj.Fnd_AnWs_BySetWs(mAnWs, pWbFm, pSetWs) Then ss.A 1: GoTo E
Dim J%
For J = 0 To jj.Siz_Ay(mAnWs) - 1
    If jj.Cpy_Ws(pWbTo, pWbFm, mAnWs(J)) Then ss.A 2: GoTo E
Next
Exit Function
R: ss.R
E: Add_Ws_ToWb = True: ss.B cSub, cMod, "pWbTo,pWbFm,pSetWs", jj.ToStr_Wb(pWbTo), jj.ToStr_Wb(pWbFm), pSetWs
End Function
#If Tst Then
Function Add_Ws_ToWb_Tst() As Boolean
Dim mXls As New Excel.Application
Dim mWbFm As Workbook: Set mWbFm = mXls.Workbooks.Add
Dim mWbTo As Workbook: Set mWbTo = mXls.Workbooks.Add
mWbFm.Sheets(3).Name = "XXXX"
If jj.Set_Ws_ByLpAp(mWbFm.Sheets(1), 1, "NmWb,NmWs,aa,bb,cc,dd", mWbFm.Name, mWbFm.Sheets(1).Name, "1", 10, #1/23/2008#, 1000) Then Stop
If jj.Set_Ws_ByLpAp(mWbFm.Sheets(2), 1, "NmWb,NmWs,aa,bb,cc,dd", mWbFm.Name, mWbFm.Sheets(2).Name, "1", 10, #1/23/2008#, 1000) Then Stop
If jj.Set_Ws_ByLpAp(mWbFm.Sheets(3), 1, "NmWb,NmWs,aa,bb,cc,dd", mWbFm.Name, mWbFm.Sheets(3).Name, "1", 10, #1/23/2008#, 1000) Then Stop
mXls.Visible = True
Stop
If Add_Ws_ToWb(mWbTo, mWbFm) Then Stop
Stop
End Function
#End If
Function Add_Cmt(pRge As Range, pCmt$, pWdt%, pHgt%) As Boolean
Const cSub$ = "Add_Cmt"
On Error GoTo R
pRge.AddComment pCmt
Dim mMaxLen%, mNLn%
With pRge.Comment.Shape
    .Width = mMaxLen * 12
    .Height = mNLn * 10 + 20
End With
R: ss.R
E: Add_Cmt = True: ss.B cSub, cMod, "pRge,pCmt,pWdt,pHgt", jj.ToStr_Rge(pRge), pCmt, pWdt, pHgt
End Function
Function Add_Rec_ByUKey_n_LpAp(oId&, pNmt$, pUKey_NmFld$, pUKey_Val$, pLnFld$, ParamArray pAp()) As Boolean
Const cSub$ = "Add_Rec_UKey_n_LpAp"
'Aim: Add or Update a record to {pNmt} with {pUKey_NmFld}, {pUKey_Val} by {pLnFld} and {mAyV} & return {oId}
'     Assume {pNmt} has an unique key of a string field {pUKey_NmFld}
'     Assume first field is the Id & AutoField field and with be returned in {oId}
Dim mRs As DAO.Recordset
Dim mSql$: mSql = jj.Fmt_Str("Select * from {0} where {1}='{2}'", pNmt, pUKey_NmFld, pUKey_Val)
If jj.Opn_Rs(mRs, mSql) Then ss.A 1: GoTo E
With mRs
    If .AbsolutePosition = -1 Then
        .AddNew
        oId = .Fields(0).Value
        .Fields(pUKey_NmFld).Value = pUKey_Val
        If jj.Set_Rs_ByLpVv(mRs, pLnFld, CVar(pAp)) Then ss.A 1: GoTo E
        .Update
    Else
        .Edit
        oId = .Fields(0).Value
        If jj.Set_Rs_ByLpVv(mRs, pLnFld, CVar(pAp)) Then ss.A 2: GoTo E
        .Update
    End If
    .Close
End With
Exit Function
R: ss.R
E: Add_Rec_ByUKey_n_LpAp = True: ss.B cSub, cMod, "pNmt,pUKey_NmFld,pUKey_Val,pLnFld,pAyV", pNmt, pUKey_NmFld, pUKey_Val, pLnFld, jj.ToStr_Vayv(CVar(pAp))
End Function
Function Add_Rec_ByUKey(oId&, pNmt$, pUKey_NmFld$, pUKey_Val) As Boolean
'Aim: Add a record to {pNmt} with {pUKey_NmFld}, {pUKey_Val}
'     Assume first field is the Id & AutoField field and with be returned in {oId}
Const cSub$ = "Add_Rec_ByUKey"
Dim mRs As DAO.Recordset
Dim mSql$: mSql = jj.Fmt_Str("Select * from {0} where {1}='{2}'", pNmt, pUKey_NmFld, pUKey_Val)
If jj.Opn_Rs(mRs, mSql) Then ss.A 1: GoTo E
With mRs
    If .AbsolutePosition = -1 Then
        .AddNew
        oId = .Fields(0).Value
        .Fields(pUKey_NmFld).Value = pUKey_Val
        .Update
    Else
        oId = .Fields(0).Value
    End If
    .Close
End With
Exit Function
R: ss.R
E: Add_Rec_ByUKey = True: ss.B cSub, cMod, "pNmt,pUKey_NmFld,pUKey_Val", pNmt, pUKey_NmFld, pUKey_Val
End Function
#If Tst Then
Function Add_Rec_ByUKey_Tst() As Boolean
Dim mID&
If jj.Add_Rec_ByUKey(mID, "xx", "bb", "1234") Then Stop
End Function
#End If
Function Add_Rec_By2Id(pNmt$, pNmFld_Pk1$, pNmFld_Pk2$, pId1&, pId2&, pLnFld$, ParamArray pAp()) As Boolean
'Aim: Add/Update a record to {pNmt} with {pLnFld} & {pAyV}
'     Assume {pNmt} has  Id fields as Pk
Const cSub$ = "Add_Rec_By2Id"
Dim mRs As DAO.Recordset
Dim mSql$: mSql = jj.Fmt_Str("Select * from {0} where {1}={2} and {3}={4}", pNmt, pNmFld_Pk1, pId1, pNmFld_Pk2, pId2)
If jj.Opn_Rs(mRs, mSql) Then ss.A 1: GoTo E
With mRs
    If .AbsolutePosition = -1 Then
        .AddNew
        .Fields(pNmFld_Pk1).Value = pId1
        .Fields(pNmFld_Pk2).Value = pId2
    Else
        .Edit
    End If
    If jj.Set_Rs_ByLpVv(mRs, pLnFld, CVar(pAp)) Then ss.A 1: GoTo E
    .Update
End With
GoTo X
R: ss.R
E: Add_Rec_By2Id = True: ss.B cSub, cMod, "pNmt,pNmFld_Pk1,pNmFld_Pk2,pId1&,pId2&,pLnFld,pAyV", pNmt$, pNmFld_Pk1$, pNmFld_Pk2$, pId1&, pId2&, pLnFld$, jj.ToStr_Vayv(CVar(pAp))
    
X:
    jj.Cls_Rs mRs
End Function
Function Add_AyAtEnd(oAy$(), pAy$(), Optional pKeepDup As Boolean = False) As Boolean
'Aim: if {pKeepDup}, Add {pAy}           to end of {oAy}
'     Else           Add those new {pAy} to end of {oAy}
'
'Note: direct assign pAy1 to oAy may have side-effect: oAy & pAy1 shared same memory space.
Const cSub$ = "Add_AyAtEnd"

Dim N1%: N1 = jj.Siz_Ay(oAy)
Dim N2%: N2 = jj.Siz_Ay(pAy)
Dim J%
If N1 = 0 Then
    If N2 = 0 Then jj.Clr_Ays oAy: Exit Function
    ReDim oAy(N2 - 1)
    For J = 0 To N2 - 1
        oAy(J) = pAy(J)
    Next
    Exit Function
End If
If N2 = 0 Then Exit Function
Dim mAy$()
If pKeepDup Then
    mAy = pAy
Else
    If jj.Ay_Subtract(mAy, pAy, oAy) Then ss.A 1: GoTo E
    N2 = jj.Siz_Ay(mAy)
    If N2 = 0 Then Exit Function
End If

ReDim Preserve oAy(N2 + N1 - 1)
For J = 0 To N2 - 1
    oAy(J + N1) = mAy(J)
Next
Exit Function
R: ss.R
E: Add_AyAtEnd = True: ss.B cSub, cMod, "oAy,pAy", jj.ToStr_Ays(oAy), jj.ToStr_Ays(pAy)
End Function
Function Add_Ay(oAy$(), pAy1$(), pAy2$(), Optional pKeepDup As Boolean = False) As Boolean
'Aim: if {pKeepDup}, Add {pAy2}           to end of {pAy1} into {oAy}
'     Else           Add those new {pAy2} to end of {pAy1} into {oAy}
'
'Note: direct assign pAy1 to oAy may have side-effect: oAy & pAy1 shared same memory space.
Const cSub$ = "Add_Ay"
Dim N1%: N1 = jj.Siz_Ay(pAy1): If N1 = 0 Then oAy = pAy2: Exit Function
Dim N2%: N2 = jj.Siz_Ay(pAy2): If N2 = 0 Then oAy = pAy1: Exit Function
Dim mAy2$()
If pKeepDup Then
    mAy2 = pAy2
Else
    If jj.Ay_Subtract(mAy2, pAy2, pAy1) Then ss.A 1: GoTo E
    N2 = jj.Siz_Ay(mAy2)
    If N2 = 0 Then oAy = pAy1: Exit Function
End If

ReDim oAy(N2 + N1 - 1)
Dim J%
For J = 0 To N1 - 1
    oAy(J) = pAy1(J)
Next
For J = 0 To N2 - 1
    oAy(J + N1) = mAy2(J)
Next
R: ss.R
E: Add_Ay = True: ss.B cSub, cMod, "oAy,pAy1,pAy2", jj.ToStr_Ays(oAy), jj.ToStr_Ays(pAy1), jj.ToStr_Ays(pAy2)
    
'Dim NNmq%: NNmq = jj.Siz_Ay(mAnq)
'Dim NNmt%: NNmt = jj.Siz_Ay(mAnt)
'If NNmt = 0 And NNmq = 0 Then
'    Dim mAntq$(): oAntq = mAntq: Exit Function
'End If
'ReDim oAntq(NNmt + NNmq - 1)
'Dim J%
'For J = 0 To NNmt - 1
'    oAntq(J) = mAnt(J)
'Next
'For I = 0 To NNmq - 1
'    oAntq(I + J) = mAnq(I)
'Next
End Function
#If Tst Then
Function Add_Ay_Tst() As Boolean
Const cSub$ = "Add_Ay_Tst"
Const N1% = 2
Const N2% = 6
Dim mAy$(), mAy1$(N1 - 1), mAy2$(N2 - 1)
Dim J%
For J = 0 To N1 - 2
    mAy1(J) = "Ay1:" & J
Next
mAy1(N1 - 1) = "Common"

For J = 0 To N2 - 2
    mAy2(J) = "Ay2:" & J
Next
mAy2(N2 - 1) = "Common"
If jj.Add_Ay(mAy, mAy1, mAy2) Then Stop
jj.Shw_Dbg cSub, cMod, "mAy,mAy1,mAy2", jj.ToStr_Ays(mAy), jj.ToStr_Ays(mAy1), jj.ToStr_Ays(mAy2)
End Function
#End If
Function Add_AyLng(oAy&(), pAy1&(), pAy2&()) As Boolean
'Aim: Add {pAy2} to end of {pAy1} into {oAy}
Const cSub$ = "Add_Ay"

Dim N1%: N1 = jj.Siz_Ay(pAy1): If N1 = 0 Then oAy = pAy2: Exit Function
Dim N2%: N2 = jj.Siz_Ay(pAy2): If N2 = 0 Then oAy = pAy1: Exit Function
ReDim Preserve oAy(N2 + N1 - 1)
Dim J%
For J = 0 To N1 - 1
    oAy(J) = pAy1(J)
Next
For J = 0 To N2 - 1
    oAy(J + N1) = pAy2(J)
Next

Exit Function
R: ss.R
E: Add_AyLng = True: ss.B cSub, cMod, "oAy,pAy1,pAy2", jj.ToStr_AyLng(oAy), jj.ToStr_AyLng(pAy1), jj.ToStr_AyLng(pAy2)
End Function
#If Tst Then
Function Add_AyLng_Tst() As Boolean
Const cSub$ = "Add_AyLng_Tst"
Const N1% = 2
Const N2% = 6
Dim mAy&(), mAy1&(N1 - 1), mAy2&(N2 - 1)
Dim J%
For J = 0 To N1 - 1
    mAy1(J) = 10 + J
Next
For J = 0 To N2 - 1
    mAy2(J) = 20 + J
Next
If jj.Add_AyLng(mAy, mAy1, mAy2) Then Stop
jj.Shw_Dbg cSub, cMod, "mAy,mAy1,mAy2", jj.ToStr_AyLng(mAy), jj.ToStr_AyLng(mAy1), jj.ToStr_AyLng(mAy2)
End Function
#End If
Function Add_Sfx_ToFfn$(pFfn$, pSfx$)
Add_Sfx_ToFfn = jj.Cut_Ext(pFfn) & pSfx & Fct.FilExt(pFfn)
End Function
Function Add_Str$(pStr$, pNew, Optional pSepChr$ = cComma)
If pStr = "" Then Add_Str = pNew: Exit Function
Add_Str = pStr & pSepChr & pNew
End Function
Function Add_Tbl_ToTbl(pNmtTar$, pNmtSrc$, pNKFld As Byte, Optional pNKFldRmv = 0) As Boolean
'Aim: Add/Upd {pNmtTar} by {pNmtSrc}.  Both has same {pNKFld} of PK.  All fields in {pNmtSrc} should all be found in {TarNmt}
'     If pNKFldRmv>0 then some record in {pNmtTar} will be remove if they does not exist in pNmtSrc having first {pNKFldRmv} as the matching keys
'     Example, Tar & Src: a,b,c, x,y,z
'              pNKFld   : 3
'              pNKFld   : 2
'              Tar: 1,1,3, ..... Src: 1,1,4, ...
'                   1,1,4, .....      1,1,5 ...
'                   1,1,5, .....      1,1,6, ...
'                 : 1,2,3, .....
'                   1,2,4, .....
'                   1,2,5, .....
'    After
'              Tar: 1,1,4
'                   1,1,5
'                   1,1,6
Const cSub$ = "Add_Tbl_ToTbl"
Dim mSqlAdd$, mSqlUpd$, mSqlDlt$
If jj.BldSql_AddUpdDlt(mSqlAdd, mSqlUpd, mSqlDlt, pNmtTar, pNmtSrc, pNKFld, pNKFldRmv) Then ss.A 1: GoTo E
jj.Shw_Sts jj.Fmt_Str("Add.Tbl_ToTbl: Adding [{0}] to [{1}] with pNKFld=[{2}] & pNKFldRmv=[{3}]", pNmtSrc, pNmtTar, pNKFld, pNKFldRmv)
If jj.Run_Sql(mSqlAdd) Then ss.A 1: GoTo E
If jj.Run_Sql(mSqlUpd) Then ss.A 1: GoTo E
If pNKFldRmv > 0 Then If jj.Run_Sql(mSqlDlt) Then ss.A 1: GoTo E
GoTo X
R: ss.R
E: Add_Tbl_ToTbl = True: ss.B cSub, cMod, "pNmtTar,pNmtSrc,pNKFld,pNKFldRmv", pNmtTar, pNmtSrc, pNKFld, pNKFldRmv
X:
    jj.Clr_Sts
End Function
#If Tst Then
Function Add_Tbl_ToTbl_Tst() As Boolean
'Create cNmtTar & cNmtSrc
Const cNPK% = 2
Const cNKFldRmv% = 0
Const cNmtTar$ = "#Add_Tbl_ToTbl_Tar", cLoFldTar$ = "aa Int, bb Int, t1 Text 10, t2 Text 10, t3 Text 10, t4 Text 10"
Const cNmtSrc$ = "#Add_Tbl_ToTbl_Src", cLoFldSrc$ = "aa Int, bb Int, t1 Text 10, t2 Text 10, t3 Text 10"

If jj.Crt_Tbl_FmLoFld(cNmtTar, cLoFldTar, cNPK) Then Stop
If jj.Crt_Tbl_FmLoFld(cNmtSrc, cLoFldSrc, cNPK) Then Stop

'Do Add data to cNmtTar & cNmtSrc
Do
    Dim J%
    Const cNRecTar% = 3 + 1
    Dim mAyRec_Tar$(cNRecTar - 1)
    mAyRec_Tar(0) = "1,0,'aa0','bb0','cc0','dd0'"
    mAyRec_Tar(1) = "1,1,'aa1','bb1','cc1','dd2'"
    mAyRec_Tar(2) = "1,2,'aa2','bb2','cc2','dd2'"
    mAyRec_Tar(3) = "1,3,'aa3','bb3','cc3','dd3'"
'    mAyRec_Tar(4) = "1,4,'aa4','bb4','cc4','dd4'"
'    mAyRec_Tar(5) = "1,5,'aa5','bb5','cc5','dd5'"
'    mAyRec_Tar(6) = "1,6,'aa6','bb6','cc6','dd6'"
'    mAyRec_Tar(7) = "1,7,'aa7','bb7','cc7','dd7'"
    Const cNRecSrc% = 7
    Dim mAyRec_Src$(cNRecSrc - 1): J = 0
    mAyRec_Src(J) = "1,1,'AA1','BB1','CC1'": J = J + 1
    mAyRec_Src(J) = "1,2,'AA2','BB2','CC2'": J = J + 1
    mAyRec_Src(J) = "1,3,'AA3','BB3','CC3'": J = J + 1
    mAyRec_Src(J) = "1,4,'AA4','BB4','CC4'": J = J + 1
    mAyRec_Src(J) = "1,5,'AA5','BB5','CC5'": J = J + 1
    mAyRec_Src(J) = "1,6,'AA6','BB6','CC6'": J = J + 1
    mAyRec_Src(J) = "1,7,'AA7','BB7','CC7'": J = J + 1
   
    Dim mSql$
    For J% = 0 To cNRecTar - 1
        mSql = jj.Fmt_Str("Insert into [{0}] values ({1})", cNmtTar, mAyRec_Tar(J))
        If jj.Run_Sql(mSql) Then Stop
    Next
    For J% = 0 To cNRecSrc - 1
        mSql = jj.Fmt_Str("Insert into [{0}] values ({1})", cNmtSrc, mAyRec_Src(J))
        If jj.Run_Sql(mSql) Then Stop
    Next
Loop Until True

If jj.Add_Tbl_ToTbl(cNmtTar, cNmtSrc, cNPK, cNKFldRmv) Then Stop
DoCmd.OpenTable cNmtTar
End Function
#End If
Function Add_Ws_ByLnWs(pWb As Workbook, pLnWs$) As Boolean
'Aim: Add {pLnWs} at end of {pWb}
Const cSub$ = "Add_Ws_ByLnWs"
On Error GoTo R
Dim mAnWs$(): mAnWs = Split(pLnWs, cComma)
Dim J%, mWsLast As Worksheet
For J = 0 To jj.Siz_Ay(mAnWs) - 1
    Set mWsLast = pWb.Sheets(pWb.Sheets.Count)
    pWb.Worksheets.Add(, mWsLast).Name = mAnWs(J)
Next
Exit Function
R: ss.R
E: Add_Ws_ByLnWs = True: ss.B cSub, cMod, "pWb,pLnWs", jj.ToStr_Wb(pWb), pLnWs
End Function
#If Tst Then
Function Add_Ws_ByLnWs_Tst() As Boolean
Const cSub$ = "Add_Ws_ByLnWs_Tst"
Const cFx$ = "c:\tmp\aa.xls"
Dim mWb As Workbook: If jj.Crt_Wb(mWb, cFx, True) Then Stop
If jj.Add_Ws_ByLnWs(mWb, "a,b,d,e,f") Then Stop
mWb.Application.Visible = True
End Function
#End If
Function Add_Ws(oWs As Worksheet, pWb As Workbook, pNmWsNew$) As Boolean
Const cSub$ = "Add_Ws"
'Aim: add and return a new ws of name <pNmWsNew> at the end of <pWb>
On Error GoTo Nxt
Set oWs = pWb.Sheets(pNmWsNew)
Exit Function
Nxt:
On Error GoTo R
Set oWs = pWb.Sheets.Add
oWs.Name = pNmWsNew
Exit Function
R: ss.R
E: Add_Ws = True: ss.B cSub, cMod, "Wb,pNmWsNew", jj.ToStr_Wb(pWb), pNmWsNew
End Function
Function Add_WsContent(pFxFm$, pWsTo As Worksheet) As Boolean
Const cSub$ = "Add_WsContent"
'Aim: Add the content of {pFxFm} at the end of {pWsTo} provided that they are same layout
'Assume: [1 Ws in pFxFm] & [Same Layout]
''[1 Ws in pFxFm] There is only one ws in {pFxFm} having the same name of the file name of {pFxFm}
''[Same Layout]       The column headings of {pFxFm} & {pWsTo} should be the same
'==Start
'Open pFxFm
Dim mWbFm As Workbook, mWsFm As Worksheet: If jj.IsSingleWsXls(pFxFm, mWbFm, mWsFm) Then ss.A 1: GoTo E
''Test if Fm & To are same column
Dim J As Byte
For J = 1 To 254
    Dim mV: mV = mWsFm.Cells(1, J).Value
    If mV = "" Then Exit Function
    If mV <> pWsTo.Cells(1, J).Value Then ss.A 2, cSub, cMod, ePrmErr, "Not same layout", , "pFxFm,Col# Not Match,Col in pFxFm,Col in pWsTo", mV, pWsTo.Cells(1, J).Value: GoTo E
Next

'Copy
Dim AdrFm$, AdrTo$
Dim mColToLast$: mColToLast = mWsFm.Cells.SpecialCells(xlCellTypeLastCell).Column
AdrFm = "A2:" & mWsFm.Cells.SpecialCells(xlCellTypeLastCell).Address
AdrTo = "A" & pWsTo.Cells.SpecialCells(xlCellTypeLastCell).Row + 1
mWsFm.Range(AdrFm).Copy
pWsTo.Range(AdrTo).PasteSpecial xlPasteValues
jj.Cls_Wb mWbFm, True
Exit Function
R: ss.R
E: Add_WsContent = True: ss.B cSub, cMod, "pFxFm,pWsTo", pFxFm, jj.ToStr_Ws(pWsTo)
End Function
#If Tst Then
Function Add_WsContent_Tst() As Boolean
Const cFfn1$ = "c:\tmp\a.xls"
Const cFfn2$ = "c:\tmp\b.xls"
Dim mWb As Workbook: Set mWb = g.gXls.Workbooks.Open(cFfn2)
Dim mWs As Worksheet: Set mWs = mWb.Sheets("sheet1")
If jj.Add_WsContent(cFfn1, mWs) Then Stop
mWb.Application.Visible = True
End Function
#End If
Function Add_WsFmCsv(oWs As Worksheet, pWbTo As Workbook, pFfnCsv$, Optional pNmWs$ = "", Optional pKillCsv As Boolean = True) As Boolean
Const cSub$ = "Add_WsFmCsv"
'Aim: Add a new ws to {pWbTo} as {pNmWs} from {pFfnCsv}.  If {pNmWs} is not given, use {pFfnCsv} as worksheet name
'Open {pFfnCsv} & Set as <mWbFm>
If Right(pFfnCsv, 4) <> ".csv" Then ss.A 1: GoTo E
Dim mWbFm As Workbook: Set mWbFm = pWbTo.Application.Workbooks.Open(pFfnCsv)

'Add a new Ws as {oWs} of name {pNmWs}
If jj.Add_Ws(oWs, pWbTo, NonBlank(pNmWs, Fct.Nam_FilNam(pFfnCsv, False))) Then ss.A 2: GoTo E

'Copy from <mWbFm.Sheet1> and Paste to <oWs>
mWbFm.Sheets(1).Cells.Copy
oWs.Cells.PasteSpecial xlPasteAll
oWs.Activate
oWs.Range("A1").Select

'Close <mWbFm>
jj.Cls_Wb mWbFm

'Kill pFfnCsv
If pKillCsv Then jj.Dlt_Fil pFfnCsv
Exit Function
R: ss.R
E: Add_WsFmCsv = True: ss.B cSub, cMod, "pWbTo,pFfnCsv,pNmWs,pKillCsv", jj.ToStr_Wb(pWbTo), pFfnCsv, pNmWs, pKillCsv
End Function
#If Tst Then
Function Add_WsFmCsv_Tst() As Boolean
End Function
#End If
