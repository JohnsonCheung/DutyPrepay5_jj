Attribute VB_Name = "xBld"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xBld"
Function Bld_TBars(pWs As Worksheet, pDclTBars$) As Boolean
'Aim: build toolbars in pWs by pDclTbars="<NmTBar>:<NmBtn1>,..<NmBtn2>; ..
Const cSub$ = "Bld_TBars"
On Error GoTo R
Dim mAnDclTBar$(): mAnDclTBar = Split(pDclTBars, ";")
Dim J%
For J = 0 To jj.Siz_Ay(mAnDclTBar) - 1
    If Bld_TBar(pWs, pDclTBars) Then ss.A 1: GoTo E
Next
R: ss.R
E: Bld_TBars = True: ss.B cSub, cMod, "pWs,pDclTBars", jj.ToStr_Ws(pWs), pDclTBars
Exit Function
End Function
Function Bld_TBar(pWs As Worksheet, pDclTBar$, Optional oTBar As MSComctlLib.ToolBar = Nothing) As Boolean
'Aim: build toolbars in pWs by pDclTbars="<NmTBar>:<NmBtn1>,..<NmBtn2>; ..
Const cSub$ = "Bld_TBar"
On Error GoTo R
Dim mNmTBar$, mLnBtn$: If jj.Brk_Str_0Or2(mNmTBar, mLnBtn, pDclTBar, ":") Then ss.A 1: GoTo E
Dim mAnBtn$(): mAnBtn = Split(mLnBtn, ",")
If Dlt_TBar(pWs, mNmTBar) Then ss.A 2: GoTo E
Dim mOLEObjs As Excel.OLEObjects
Set oTBar = mOLEObjs.Add("MSComctlLib.ToolBar")
Dim J%
For J = 0 To jj.Siz_Ay(mAnBtn) - 1
    oTBar.Buttons.Add , , mAnBtn(J)
Next
Exit Function
R: ss.R
E: Bld_TBar = True: ss.B cSub, cMod, "pWs,pDclTBars", jj.ToStr_Ws(pWs), pDclTBar
Exit Function
End Function
#If Tst Then
Function Bld_TBars_Tst() As Boolean
Dim mWb As Workbook: Set mWb = Excel.Application.Workbooks("MetaTx.xls")
Dim mWs As Worksheet: Set mWs = mWb.Sheets("Tx")
If Bld_TBars(mWs, "aa:df,sdf;bb:slkdf,dsf") Then Stop: GoTo E
Exit Function
E: Bld_TBars_Tst = True
End Function
#End If
Function Bld_SqlSel$(pSel$, pFm$, Optional pWhere$ = "", Optional pOrdBy$ = "")
Const cSqlSel$ = "Select {0} from {1}{2}{3}"
Bld_SqlSel = jj.Fmt_Str(cSqlSel, pSel, jj.Q_S(pFm, "[]"), jj.Cv_Where(pWhere), jj.Cv_OrdBy(pOrdBy))
End Function
Function Bld_Struct_ForTy_Import$(pItm$, pMaxTy As Byte)
'Aim: Build {mR} of a Import table [>{pItm}] having {pMaxTy}>0
'     Assume in there is import table named as [>{pItm}] and pMaxTy=2, the return {mR} will be
'       Nm<pItm>, and,
'       Nm<pItm>Ty1, Nm<pItm>Ty1x, Nm<pItm>Ty1xx, Nm<pItm>Ty1xxx, and,
'       Nm<pItm>Ty2, Nm<pItm>Ty2x, Nm<pItm>Ty2xx, Nm<pItm>Ty2xxx.
'     Eg <pItm>=Tbl
'       NmTbl, and,
'       NmTblTy1, NmTblTy1x, NmTblTy1xx, NmTblTy1xxx, and,
'       NmTblTy2, NmTblTy2x, NmTblTy2xx, NmTblTy2xxx.
Const cSub$ = "Bld_Struct_ForTy_Import"
Dim mR$
If pMaxTy = 0 Then ss.A 1, "pMaxTy must >0", , "pItm,pMaxTy", pItm, pMaxTy: GoTo E
On Error GoTo R
mR = "Nm" & pItm
Dim J%
For J = 1 To pMaxTy
    mR = mR & jj.Fmt_Str(",NmTy{0}{1},NmTy{0}{1}x,NmTy{0}{1}xx,NmTy{0}{1}xxx", pItm, J)
Next
Bld_Struct_ForTy_Import = mR
Exit Function
R: ss.R
E: ss.B cSub, cMod, "pItm,pMaxTy", pItm, pMaxTy: GoTo E
End Function
Function Bld_Struct_ForTy_Import_Tst() As Boolean
Debug.Print jj.Bld_Struct_ForTy_Import("Tbl", 2)
End Function
Function Bld_Struct_ForTy$(pItm$, Optional pN As Byte = 0, Optional pX$ = "")
'Aim: Build {mR} (a list of field (may with field type and len for table creation) from {pItm}, {pN}, {pX}
'           pN is 0 to 5 (MaxTy)
'           pX is "", "x", .., "xxx"
'       Ty Tables:   name = [$<pItm>Ty]     Example, $TblTy for each record in $Tbl.  3 fields: Tbl, TblTy1, TblTy2
'       Ty1 Tables:  name = [$<pItm>Ty1]    Example, $TblTy1.                         4 fields: TblTy1,    NmTblTy1,    TblTy1x,   DesTblTy1
'                           [$<pItm>Ty1x]   Example, $TblTy1x                         4 fields: TblTy1x,   NmTblTy1x,   TblTy1xx,  DesTblTy1
'                           [$<pItm>Ty1xx]  Example, $TblTy1xx                        4 fields: TblTy1xx,  NmTblTy1xx,  TblTy1xxx, DesTblTy1
'                           [$<pItm>Ty1xxx] Example, $TblTy1xxx                       3 fields: TblTy1xxx, NmTblTy1xxx,            DesTblTy1
'       Ty2 Tables:  name = [$<pItm>Ty2]    Example, $TblTy2.                         4 fields: TblTy2,    NmTblTy2,    TblTy2x,   DesTblTy2
'                           [$<pItm>Ty2x]   Example, $TblTy2x                         4 fields: TblTy2x,   NmTblTy2x,   TblTy2xx,  DesTblTy2
'                           [$<pItm>Ty2xx]  Example, $TblTy2xx                        4 fields: TblTy2xx,  NmTblTy2xx,  TblTy2xxx, DesTblTy2
'                           [$<pItm>Ty2xxx] Example, $TblTy2xxx                       3 fields: TblTy2xxx, NmTblTy2xxx,            DesTblTy2
'Note: pN=0, pX will be "1", .., "5" (which means MaxTy), Ty Table will be return
Const cSub$ = "Bld_Struct_ForTy"
On Error GoTo R
Dim mR$
If pN = 0 Then
    If 1 > Val(pX) Or Val(pX) > 5 Then ss.A 1, "pN=0, pX must be '1', ..,'5', which means MaxTy", ePrmErr: GoTo E
ElseIf pN > 5 Then
    ss.A 2, "pN should 0-5", ePrmErr: GoTo E
Else
    If pX <> "" And pX <> "x" And pX <> "xx" And pX <> "xxx" Then ss.A 1, "If pN between 1-5, pX must be '','x',..'xxx'", ePrmErr: GoTo E
End If
', Optional pForCrt As Boolean = False
'If pForCrt Then Exit Function
If pN = 0 Then
    Dim J%
    mR = pItm
    For J = 1 To CByte(pX)
        mR = mR & ",Ty" & pItm & J
    Next
    Exit Function
End If
If pX = "xxx" Then
    mR = jj.Fmt_Str("Ty{0}{1}{2},NmTy{0}{1}{2},DesTy{0}{1}{2}", pItm, pN, pX)
Else
    mR = jj.Fmt_Str("Ty{0}{1}{2},NmTy{0}{1}{2},Ty{0}{1}{2}x,DesTy{0}{1}{2}", pItm, pN, pX)
End If
Bld_Struct_ForTy = mR
Exit Function
R: ss.R
E: ss.B cSub, cMod, "pItm,pN,pX", pItm, pN, pX: GoTo E
End Function
#If Tst Then
Function Bld_Struct_ForTy_Tst() As Boolean
Debug.Print jj.Bld_Struct_ForTy_Import("Tbl", 3)
Debug.Print jj.Bld_Struct_ForTy("Tbl", 0, "3")
Debug.Print jj.Bld_Struct_ForTy("Tbl", 1, "")
Debug.Print jj.Bld_Struct_ForTy("Tbl", 1, "x")
Debug.Print jj.Bld_Struct_ForTy("Tbl", 1, "xx")
Debug.Print jj.Bld_Struct_ForTy("Tbl", 1, "xxx")
Debug.Print jj.Bld_Struct_ForTy("Tbl", 2, "")
Debug.Print jj.Bld_Struct_ForTy("Tbl", 2, "x")
Debug.Print jj.Bld_Struct_ForTy("Tbl", 2, "xx")
Debug.Print jj.Bld_Struct_ForTy("Tbl", 2, "xxx")
Debug.Print jj.Bld_Struct_ForTy("Tbl", 3, "")
Debug.Print jj.Bld_Struct_ForTy("Tbl", 3, "x")
Debug.Print jj.Bld_Struct_ForTy("Tbl", 3, "xx")
Debug.Print jj.Bld_Struct_ForTy("Tbl", 3, "xxx")
Shw_DbgWin
End Function
#End If
Function Bld_Where(oWhere$, pLp$, pVayv) As Boolean
'Aim:
Const cSub$ = "Bld_Where"
If pLp = "" Then oWhere = "": Exit Function
Dim mAn$(): mAn = Split(pLp$, cComma)
Dim mAv$(): mAv = pVayv
Dim N1%: N1 = jj.Siz_Ay(mAn)
Dim N2%: N2 = jj.Siz_Ay(mAv)
If N1 <> N2 Then ss.A 1, "Count in pLp & pVayv mismatch", "Cnt in pLp, Cnt in pVayv", , N1, N2: GoTo E
Dim J%: For J = 0 To N1 - 1
    Dim mA$: mA = mAn(J) & " in (" & mAv(J) & ")"
    oWhere = jj.Add_Str(oWhere, mA, " and ")
Next
If oWhere <> "" Then oWhere = " Where " & oWhere
Exit Function
E: Bld_Where = True: ss.B cSub, cMod, "pLp,pVayv", pLp, jj.ToStr_Ays(mAv)
End Function
#If Tst Then
Function Bld_Where_Tst() As Boolean
Dim mWhere$, mAv$(2)
mAv(0) = "11,22,33"
mAv(1) = "44,55,66"
mAv(2) = "77,88"
If jj.Bld_Where(mWhere, "aa,bb,cc", mAv) Then Stop
Debug.Print mWhere
End Function
#End If
Function Bld_LExpr(oLExpr$, pNmFld$, pTypSim As eTypSim, pVraw$, Optional pIsOpt As Boolean = False) As Boolean
'Aim: Build a sql condition {oLExpr}
'Prm: {pVraw}   [x] [%x-x] [x,x,x] [>x] [>=x] [<x] [<=x] [*x] [x*] [*x*] [!%x-x] [!x,x,x] [!*x] [!x*] [!*x*] [!x] (16)
'               Eq  Rge    Lst     Gt   Ge    Lt   Le    ----- Lik ----- NRge    NLst     ------ NLik ------ Ne   (12)
Const cSub$ = "Bld_LExpr"
On Error GoTo R
If pIsOpt Then If Trim(pVraw) = "" Then ss.A 1, "Not optional prm, but pVraw is empty": GoTo E
Select Case pTypSim
Case eTypSim_Num, eTypSim_Bool, eTypSim_Str, eTypSim_Dte
Case Else
     ss.A 2, "Only TypSim: N,B,S,D will handle": GoTo E
End Select
If pNmFld = "" Then ss.A 3, "pNmfld cannot be nothing": GoTo E

'Find [OpTyp,V1,V2] by {pVraw}
Dim V1$, V2$, OpTyp As eOpTyp
''(5) [Test C2 first]  [>=] [<=] [!%] [!*x] [!*x*]
Dim c2$: c2 = Left(pVraw, 2)
Select Case c2
Case ">=": OpTyp = eGe: V1 = mID(pVraw, 3)
Case "<=": OpTyp = Ele: V1 = mID(pVraw, 3)
Case "!%": OpTyp = eNRge: If jj.Brk_Str_Both(V1, V2, mID(pVraw, 3), "-") Then ss.A 4, "no - in !%xx-xx": GoTo E
Case "!*": OpTyp = eNLik: V1 = mID(pVraw, 2)
Case Else
    ''(9) [Test C1 next]   [%]  [>]  [<]  [*x]  [*x*]  [!x:x:x] [!x*] [!x]
    Dim c1$: c1 = Left(c2, 1)
    Select Case c1
    Case "%": OpTyp = eRge: If jj.Brk_Str_Both(V1, V2, mID(pVraw, 2), "-") Then ss.A 5, "no - in %xx-xx": GoTo E
    Case ">": OpTyp = eGt: V1 = mID(pVraw, 2)
    Case "<": OpTyp = eLt: V1 = mID(pVraw, 2)
    Case "*": OpTyp = eLik: V1 = pVraw
    Case "!": V1 = mID(pVraw, 2) '[!x,x,x] [!x*] [!x]
                If InStr(V1, cComma) > 0 Then
                    OpTyp = eNLst
                Else
                    If InStr(V1, "*") > 0 Then
                        OpTyp = eNLik
                    Else
                        OpTyp = eNe
                    End If
                End If
    Case Else
        ''(1) [Test x,x,x]
        Dim Pos As Byte: Pos = InStr(pVraw, cComma)
        If Pos > 0 Then
            OpTyp = eLst: V1 = pVraw
        Else
            ''(1) [x]
            OpTyp = eEq: V1 = pVraw
        End If
    End Select
End Select

'[Find Q]
''All Op: Eq , Rge, Lst, Gt, Ge, Lt, Le, Lik, NRge, NLst, NLik, Ne
Dim Q$
Select Case pTypSim
''[Bool: Op=(Eq,Ne)]
Case eTypSim_Bool
    Select Case OpTyp
    Case eOpTyp.eEq, eOpTyp.eNe
    Case Else:  ss.A 6, "Boolean data only allow = or <>", , "OpTyp", OpTyp: GoTo E
    End Select
    If V1 <> "True" And V1 <> "False" Then ss.A 7, "Boolean data allow value True or False", , "V1", V1: GoTo E
Case eTypSim_Dte
''[Dte: Op=(!Lik,NLik)]
    Select Case OpTyp
    Case eOpTyp.eLik, eOpTyp.eNLik:  ss.A 8, "Date data not allow Like or not like", , "OpTyp", OpTyp: GoTo E
    Case Else
    End Select
    Q = "#"
Case eTypSim_Str
    Q = cQSng
End Select

'[Reject Some Value for some TypSim]
''All Op: Eq , Rge, Lst, Gt, Ge, Lt, Le, Lik, NRge, NLst, NLik, Ne
Dim Ay$(), J%
Select Case OpTyp
Case eOpTyp.eEq, eOpTyp.eGt, eOpTyp.eGe, eOpTyp.eLt, eOpTyp.Ele, eOpTyp.eLik, eOpTyp.eNLik, eOpTyp.eNe
    If jj.Cv_Vraw2Val(V1, V1, pTypSim) Then ss.A 9, "Field 1 has invalid value": GoTo E
Case eOpTyp.eRge, eOpTyp.eNRge
    Dim mV1: If jj.Cv_Vraw2Val(mV1, V1, pTypSim) Then ss.A 10, "Field 1 has invalid value", , "V1", V1: GoTo E
    Dim mV2: If jj.Cv_Vraw2Val(mV2, V2, pTypSim) Then ss.A 11, "Field 2 has invalid value", , "V1", V2: GoTo E
    If mV1 > mV2 Then ss.A 12, "Field 2 > Field 1 for Between or Not Between", , "V1,V2", V1, V2: GoTo E
Case eOpTyp.eLst, eOpTyp.eNLst
    Ay = Split(V1, cComma)
    For J = LBound(Ay) To UBound(Ay)
        If jj.Cv_Vraw2Val(Ay(J), Ay(J), pTypSim) Then ss.A 13, "Some field of list data has invalid value", "Ay(J)", Ay(J): GoTo E
    Next
Case Else
    If jj.Cv_Vraw2Val(Ay(J), Ay(J), pTypSim) Then ss.A 1: GoTo E
End Select

'[Normalize V1,V2 if Q<>'']
If Q <> "" Then
    Select Case OpTyp
    Case eOpTyp.eEq, eOpTyp.eGt, eOpTyp.eGe, eOpTyp.eLt, eOpTyp.Ele, eOpTyp.eLik, eOpTyp.eNLik, eOpTyp.eNe
        V1 = Q & V1 & Q
    Case eOpTyp.eRge, eOpTyp.eNRge
        V1 = Q & V1 & Q: V2 = Q & V2 & Q
    Case eOpTyp.eLst, eOpTyp.eNLst
        Ay = Split(V1, cComma)
        For J = LBound(Ay) To UBound(Ay)
            Ay(J) = Q & Ay(J) & Q
        Next
        V1 = Join(Ay, cComma)
    Case Else
         ss.A 14, "Unexpected OpTyp for data need quote", , "Q,OpTyp", Q, OpTyp: GoTo E
    End Select
End If

X:
Static AyOpFmtStr$(1 To 12)
If AyOpFmtStr(1) = "" Then
    AyOpFmtStr(eOpTyp.eLik) = "({0} like {1})"
    AyOpFmtStr(eOpTyp.eEq) = "({0}={1})"
    AyOpFmtStr(eOpTyp.eGe) = "({0}>={1})"
    AyOpFmtStr(eOpTyp.eGt) = "({0}>{1})"
    AyOpFmtStr(eOpTyp.Ele) = "({0}<={1})"
    AyOpFmtStr(eOpTyp.eLst) = "({0} in ({1}))"
    AyOpFmtStr(eOpTyp.eLt) = "({0}<{1})"
    AyOpFmtStr(eOpTyp.eNLik) = "({0} not like {1})"
    AyOpFmtStr(eOpTyp.eNLst) = "({0} not in ({1}))"
    AyOpFmtStr(eOpTyp.eNRge) = "({1}>{0} or {0}>{2})"
    AyOpFmtStr(eOpTyp.eRge) = "({0} between {1} and {2})"
    AyOpFmtStr(eOpTyp.eNe) = "({0}<>{1})"
End If
oLExpr = jj.Fmt_Str(AyOpFmtStr(OpTyp), pNmFld, V1, V2)
Exit Function
R: ss.R
E: Bld_LExpr = True: ss.B cSub, cMod, "pIsOpt,pNmfld,pTypSim,pVraw", pIsOpt, pNmFld, pTypSim, pVraw
End Function
Function Bld_LExpr_ByAyNm2V(oLExpr$, pAyNm2V() As tNm2V, Optional pAlwNull As Boolean = False) As Boolean
'Aim: Build {oLExpr} by NEW value in {pAyNm2V}.  Any Null triggers error.
Const cSub$ = "Bld_LExpr_ByAyNm2V"
oLExpr = ""
Dim J%, mIsEq As Boolean
For J = 0 To jj.Siz_An2V(pAyNm2V) - 1
    With pAyNm2V(J)
        If VarType(.NewV) = vbNull Then
            If Not pAlwNull Then ss.A 1, "The one of the element of .NewV in pAyNm2V is Null", , "The Ele,J", pAyNm2V(J).Nm, J: GoTo E
            oLExpr = jj.Add_Str(oLExpr, jj.Q_S(.Nm, "IsNull(*)"), " and ")
        Else
            oLExpr = jj.Add_Str(oLExpr, .Nm & "=" & jj.Q_V(.NewV), " and ")
        End If
    End With
Next
Exit Function
E: Bld_LExpr_ByAyNm2V = True: ss.B cSub, cMod, "pAyNm2V", jj.ToStr_AyNm2V(pAyNm2V)
End Function
Function Bld_LExpr_InFrm(oLExpr$, pFrm As Form, pLmPk$) As Boolean
'Aim: Build {oLExpr} by OldValue of {pLmPk$} in {pFrm} with optional to replace the variable name by {pLnNew}
Const cSub$ = "Bld_LExpr_InFrm"
Dim mAyNm2V() As tNm2V: If jj.Fnd_An2V_ByFrm(mAyNm2V, pFrm, pLmPk) Then ss.A 1: GoTo E
Bld_LExpr_InFrm = jj.Bld_LExpr_ByAyNm2V(oLExpr, mAyNm2V)
Exit Function
R: ss.R
E: Bld_LExpr_InFrm = True: ss.B cSub, cMod, ""
End Function
Function Bld_LExpr_ByLpVv(oLExpr$, pLp$, pVayv) As Boolean
'Aim: Build condition {oLExpr} by {pLn} and a variant which is array of variant of value {pVayv}.  {Vayv} mean Variant that storing array of variant. {Val} means value
Const cSub$ = "Bld_LExpr_ByLpVv"
If VarType(pVayv) <> vbArray + vbVariant Then ss.A 1, "VarType of pVayv must be Array+Var", , "VarType(pVayv)", VarType(pVayv): GoTo E
oLExpr = ""
Dim mAn$(): mAn = Split(pLp, cComma)
Dim mAyV(): mAyV = pVayv
Dim N1%: N1 = jj.Siz_Ay(mAn)
Dim N2%: N2 = jj.Siz_Ay(mAyV)
If N1 <> N2 Then ss.A 1, "Cnt in pLn & pV() not match", , "N1,N2", N1, N2: GoTo E
Dim J%: For J = 0 To N1 - 1
    Dim mA$: If jj.Join_NmV(mA, mAn(J), mAyV(J)) Then ss.A 2: GoTo E
    oLExpr = jj.Add_Str(oLExpr, mA, " and ")
Next
Exit Function
R: ss.R
E: Bld_LExpr_ByLpVv = True: ss.B cSub, cMod, "pLn,pV", pLp, jj.ToStr_Vayv(pVayv)
End Function
Function Bld_LExpr_ByLpAp(oLExpr$, pLp$, ParamArray pAp()) As Boolean
'Aim: Build condition {oLExpr} by {pLn} and values in {pAp()}
Const cSub$ = "Bld_LExpr_ByLpAp"
If jj.Bld_LExpr_ByLpVv(oLExpr, pLp, CVar(pAp)) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: Bld_LExpr_ByLpAp = True: ss.B cSub, cMod, "pLp,pAp", pLp, jj.ToStr_Vayv(CVar(pAp))
End Function
Function Bld_LExpr_ByLpVv_Tst() As Boolean
Const cSub$ = "Bld_LExpr_ByLpAp_Tst"
Dim mCndn$, mLn$, mV1, mV2, mV3
Dim mRslt As Boolean, mCase As Byte: mCase = 1
Select Case mCase
Case 1
    mLn$ = "ss,nn,dd"
    mV1 = "aa"
    mV2 = 11
    mV3 = #2/1/2007#
End Select
mRslt = jj.Bld_LExpr_ByLpAp(mCndn, mLn, mV1, mV2, mV3)
jj.Shw_Dbg cSub, cMod, "mRslt,mCndn,mLn,mV1,mV2,mV3", mRslt, mCndn, mLn, mV1, mV2, mV3
End Function
#If Tst Then
Function Bld_LExpr_Tst() As Boolean
Const cSub$ = "Bld_LExpr_Tst"
Dim J%, N%, mCndn$
N = 20
ReDim mN$(N)
J = -1
J = J + 1: mN(J) = "!"
J = J + 1: mN(J) = "!123"
J = J + 1: mN(J) = "!123,124"
J = J + 1: mN(J) = "!%123-124"
J = J + 1: mN(J) = "!123-124"
J = J + 1: mN(J) = "!*123"
J = J + 1: mN(J) = "!*123-1234"
J = J + 1: mN(J) = "123:124"
J = J + 1: mN(J) = "123-124"
J = J + 1: mN(J) = "%123:124"
J = J + 1: mN(J) = "%123-124"
J = J + 1: mN(J) = "%125-124"
J = J + 1: mN(J) = "123,124"
J = J + 1: mN(J) = "124"
J = J + 1: mN(J) = "*123"
J = J + 1: mN(J) = "123*"
J = J + 1: mN(J) = "*123*"
J = J + 1: mN(J) = "!*123"
J = J + 1: mN(J) = "!123*"
J = J + 1: mN(J) = "!*123"
jj.Shw_DbgWin
jj.Set_Silent
For J = 0 To N
    If mN(J) = "" Then Exit For
    Debug.Print J, mN(J),
    If jj.Bld_LExpr(mCndn, "abc", eTypSim_Str, mN(J)) Then
        Debug.Print "<-- Error"
    Else
        Debug.Print "<-- "; mCndn
    End If
Next
For J = 0 To N
    If mN(J) = "" Then Exit For
    Debug.Print J, mN(J),
    If jj.Bld_LExpr(mCndn, "abc", eTypSim_Num, mN(J)) Then
        Debug.Print "<-- Error"
    Else
        Debug.Print "<-- "; mCndn
    End If
Next
jj.Shw_DbgWin
jj.Set_Silent_Rst
Exit Function
E: Bld_LExpr_Tst = True
End Function
#End If
Function Bld_Dtf(pFfnDtf$, pSql$, pIP$ _
    , Optional pLib$ = "RBPCSF" _
    , Optional pIsByXls As Boolean = False _
    , Optional pRun As Boolean = False _
    , Optional oNRec& = 0 _
    ) As Boolean
Const cSub$ = "Bld_Dtf"
'Aim: Build a file [pFfnDtf]  by [{pSql}, {pIP}, {pLib}] with optional to run it.
'     which will download data to [mFfnTar] with Fdf in same directory as pFfnDtf.  If no data is download empty Txt or empty Xls will be created according to FDF
'     [mFfnFdf] = Ffnn(pFfnDtf).Fdf
'     [mFfnTar] = Ffnn(pFfnDtf).Txt (or .xls)
'     ResStr @ modResStr.DtfTp()
If Right(pFfnDtf, 4) <> ".Dtf" Then ss.A 1, "pFfnDtf must end with .dtf": GoTo E

Dim mDtfTp$, mFfnTar$
Do
    'Find [mDtfTp] & create [mFfnDtf] by following following macro string
    '               Txt     Xls
    '{IP}
    '{LIB}
    '{Sql}
    '{ConvTyp}      0       1
    '{FfnFDF}
    '{FfnTar}
    '{PCFilTyp}     1       16
    '{SavFDF}       1       1
    Dim mConvTp$, mPCFilTyp$, mSavFDF$, mFfnFdf$
    Dim mFfnn$: mFfnn = jj.Cut_Ext(pFfnDtf)
    mFfnFdf$ = mFfnn & ".Fdf"
    mSavFDF = "1"
    If pIsByXls Then
        mConvTp = "4" ' "1"
        mFfnTar = mFfnn & ".xls"
        mPCFilTyp = "16"
    Else
        mConvTp = "0"
        mFfnTar = mFfnn$ & ".Txt"
        mPCFilTyp = "1"
    End If
    If jj.Dlt_Fil(mFfnTar) Then ss.A 1: GoTo E
    Bld_zFndDtfTp mDtfTp, pIP$, pLib$, pSql$, mConvTp$, mFfnFdf, mFfnTar, mPCFilTyp$, mSavFDF$
    
    'Create {pFnnDtf}
    If jj.Exp_Str_ToFfn(mDtfTp, pFfnDtf, True) Then ss.A 2: GoTo E
Loop Until True
If pRun Then If jj.Run_Dtf(pFfnDtf, mFfnTar, oNRec) Then ss.A 3: GoTo E
Exit Function
R: ss.R
E: Bld_Dtf = True: ss.B cSub, cMod, "pFfnDtf,pSql,pIP,pLib,pIsByXls,pRun,oNRec", pFfnDtf, pSql, pIP, pLib, pIsByXls, pRun, oNRec
End Function
#If Tst Then
Function Bld_Dtf_Tst() As Boolean
Dim mNRec&
'If Dtf("Dtf", "AVM_Xls", "Select * from AVM", "192.168.103.14", , , True, True, True, mNRec) Then Stop Else Debug.Print mNRec
'If Dtf("c:\Tmp\IIC.Dtf", "Select * from IIC where iclas='ux'", "192.168.103.13", "BPCSF", False, True, mNRec) Then Stop Else Debug.Print mNRec
If jj.Bld_Dtf("c:\Tmp\IIC.Dtf", "Select * from IIC where iclas='07'", "192.168.103.13", "BPCSF", True, True, mNRec) Then Stop Else Debug.Print mNRec
Debug.Print mNRec
End Function
#End If
Function Bld_OdbcTwoQry_BySqlDsn(pSelSql$, pDsn$, pIsCrtTbl As Boolean, pNmtTar$ _
        , Optional pInFbTar$ = "" _
        , Optional pNmq1$ = "OdbcTwoQry_BySqlDsn_Q1" _
        , Optional pNmq2$ = "OdbcTwoQry_BySqlDsn_Q2" _
        , Optional pRun As Boolean = False) As Boolean
'Aim: Create 2 queries having {pNmq1} which will either be create or append the {pNmtTar} in {pInFbTar}
'       {pNmq2} will be a Dsn select query using {pSelSql} & {pDsn} & [jj.Crt_Qry_ByDsn]
'       {pNmq1} is either a create/append query depend on {pIsCrtTbl}:
'                   from table      {pNmq2}
'                   Target Table    {pNmtTar} in {pFb}
'                   Select fields   either * or {mLnFld_Tar} which will convert any field having [yymd_*] in {pSelSql} to
'                                   Cdate(IIf(yymd_{0}=0,0,IIf(yymd_{0}=99999999,'9999/12/31',format(yymd_{0},'0000\/00\/00')))) as {0}", Mid(mAy(I), 6)
Const cSub$ = "Bld_OdbcTwoQry_BySqlDsn"
Const cSql_Crt$ = "Select {0} into {1} from {2}"
Const cSql_App$ = "Insert into {1} Select {0} from {2}"

'Build Query 2
Dim mDb As DAO.Database: If jj.Cv_Db_FmFb(mDb, pInFbTar) Then ss.A 1: GoTo E
If jj.SysCfg_IsLclMd Then
    If jj.Crt_Qry(pNmq2, pSelSql) Then ss.A 2: GoTo E
Else
    If jj.Crt_Qry_ByDSN(pNmq2, pSelSql, pDsn, True) Then ss.A 3: GoTo E
End If

'Build Query 1
''Fnd [mLnFld_Tar] which will convert yymd_* as date from list of Build
Dim mLnFld_Tar$: mLnFld_Tar = "*"
If InStr(pSelSql, "as yymd_") > 0 Then If jj.Fnd_LnFld_ByNmq(mLnFld_Tar, pNmq2) Then ss.A 4: GoTo E

''Build Query 1
Dim mNmtTar_wInFbTar$
If pInFbTar = "" Then
    mNmtTar_wInFbTar = pNmtTar
Else
    mNmtTar_wInFbTar = pNmtTar & " in '" & pInFbTar & cQSng
End If
Dim mSql$: mSql = jj.Fmt_Str(IIf(pIsCrtTbl, cSql_Crt, cSql_App), mLnFld_Tar, mNmtTar_wInFbTar, pNmq2) 'Const cSql_Crt$ = "Select {0} into {1} from {2}"
If jj.Crt_Qry(pNmq1, mSql) Then ss.A 5: GoTo E

If pRun Then If jj.Run_Sql(mSql) Then ss.A 6: GoTo E
If Not jj.SysCfg_IsDbgOdbc Then
    jj.Dlt_Qry pNmq1, mDb
    jj.Dlt_Qry pNmq2, mDb
End If
If pInFbTar = "" Then jj.Cls_Db mDb
Exit Function
R: ss.R
E: Bld_OdbcTwoQry_BySqlDsn = True: ss.B cSub, cMod, "pSelSql,pDsn,pIsCrtTbl,pNmtTar,pInFbTar,pNmq1,pNmq2,pRun", pSelSql, pDsn, pIsCrtTbl, pNmtTar, pInFbTar, pNmq1, pNmq2, pRun
X:
    If pInFbTar = "" Then jj.Cls_Db mDb
End Function
#If Tst Then
Function Bld_OdbcTwoQry_BySqlDsn_Tst() As Boolean
Const cSub$ = "Bld_OdbcTwoQry_BySqlDsn_Tst"
Dim mCase%: mCase = 1
Dim mResult As Boolean
Dim mNmq1$, mNmq2$, mNmtTar$, mDsn$, mSql$, mFbTar$, mRun As Boolean
mFbTar = "c:\aa.mdb"
Select Case mCase
Case 1
    mNmq1 = "qryMPS_1"
    mNmq2 = "qryMPS_2"
    mNmtTar = "tmpFSO"
    mDsn$ = "FEPROD_RBPCSF"
    mSql$ = "Select LODTE AS yymd_EnterDte, LRDTE AS yymd_DueDte, ECL.* from ECL WHERE LPROD LIKE '07%' AND LQORD>LQSHP"
    mRun = True
    mResult = jj.Bld_OdbcTwoQry_BySqlDsn(mSql, mDsn, True, mNmtTar, mFbTar, mNmq1, mNmq2, mRun)
Case 2
End Select
jj.Shw_Dbg cSub, cMod, "Result, mNmq1, mNmq2, mFbTar, mNmtTar, mDsn, mSql, mRun", mResult, mNmq1, mNmq2, mFbTar, mNmtTar, mDsn, mSql, mRun
Debug.Print "----The content of the TwoQueries --"
Debug.Print jj.ToStr_LpAp(vbLf, "mNmq1.Sql,mNmq2.Sql", CurrentDb.QueryDefs(mNmq1).Sql, CurrentDb.QueryDefs(mNmq2).Sql)
End Function
#End If
Function Bld_OdbcTwoQry_BySqlDtf(pSelSql$, pIP$, pLib$, pIsCrtTbl As Boolean, pNmtTar$, pNmtTar_DtfTxt$, pNmq1$, pNmq2$ _
    , Optional pIsByXls As Boolean = False _
    , Optional pFbTar$ = "" _
    , Optional pRun As Boolean = False) As Boolean
'Aim: Create 2 queries & {pRun} {pNmq1} optionally to {pIsCrtTbl} (Create/Append) {pNmtTar} in {pFbTar}
' {pNmq2} will be a selection query of "Select * from {pNmtTar_DtfTxt}",
'         where {pNmtTar_DtfTxt} is an imported table from a text file [mFfnDtfTxt],
'         where [mFfnDtfTxt] is downloaded by {pSelSql$, pIP$, pLib$} by [mFfn_Dtf]
'         where [mFfn_Dtf] & [mFfn_DtfTxt] will be located in [cDir] having name {pNmtTar} of ext .Dtf or .Txt
'         where [cDir] = {DirTmp}Bld_OdbcTwoQry_BySqlDtf
' {pNmq1} is either a create table / append table query depend on {pIsCrtTbl}:
'         from table      {pNmq2}
'         Target Table    {pNmtTar} in {pFbTar}
'         Select fields   either * or {mLnFld_Tar} which will convert any field having [yymd_*] in {pSelSql} to
'                         Cdate(IIf(yymd_{0}=0,0,IIf(yymd_{0}=99999999,'9999/12/31',format(yymd_{0},'0000\/00\/00')))) as {0}", Mid(mAy(I), 6)
Const cSub$ = "Bld_OdbcTwoQry_BySqlDtf"
If jj.Crt_Tbl_FmDTF_Sql(pIP, pSelSql, pNmtTar_DtfTxt, pFbTar, pLib, pIsByXls) Then ss.A 1: GoTo E
Dim mDb As DAO.Database: If jj.Cv_Db_FmFb(mDb, pFbTar) Then ss.A 2: GoTo E
If jj.Crt_Qry(pNmq2, "Select * from " & pNmtTar_DtfTxt, mDb) Then ss.A 3: GoTo E

'Fnd [mLnFld_Tar] which will convert yymd_* as date from list of Build
Dim mLnFld_Tar$: If jj.Fnd_LnFld_ByNmq(mLnFld_Tar, pNmq2, , , mDb) Then ss.A 4: GoTo E

'Build Query 1
Const cSql_Crt$ = "Select {0} into {1} from {2}"
Const cSql_App$ = "Insert into {1} Select {0} from {2}"
Dim mSql$: mSql = jj.Fmt_Str(IIf(pIsCrtTbl, cSql_Crt, cSql_App), mLnFld_Tar, pNmtTar, pNmq2)
If jj.Crt_Qry(pNmq1, mSql, mDb) Then ss.A 5: GoTo E
'Optionally {pRun} {pNmq1} to {pIsCrtTbl} (Create/Append) {pNmtTar} in {pFbTar}
If pRun Then If jj.Run_Sql_ByDbExec(mSql, mDb) Then ss.A 6: GoTo E
GoTo X
R: ss.R
E: Bld_OdbcTwoQry_BySqlDtf = True: ss.B cSub, cMod, "pSelSql,pIP,pLib,pIsCrtTbl,pNmtTar,pNmtTar_DtfTxt,pNmq1,pNmq2,pIsByXls,pFbTar,pRun", pSelSql, pIP, pLib, pIsCrtTbl, pNmtTar, pNmtTar_DtfTxt, pNmq1, pNmq2, pIsByXls, pFbTar, pRun
X:
    If pFbTar <> "" Then jj.Cls_Db mDb
End Function
#If Tst Then
Function Bld_OdbcTwoQry_BySqlDtf_Tst() As Boolean
Const cSub$ = "Bld_OdbcTwoQry_BySqlDtf_Tst"
Dim mCase%: mCase = 1
Dim mResult As Boolean
Dim mNmq1$, mNmq2$, mNmtTar$, mIP$, mLib$, mSql$, mNmtTar_DtfTxt$, mIsByXls As Boolean
Dim mFbTar$: mFbTar = "c:\aa.mdb"
Dim mRun As Boolean: mRun = True
Select Case mCase
Case 1
    mNmq1 = "qryECL_1"
    mNmq2 = "qryECL_2"
    mNmtTar = "tmpECL"
    mIP$ = "192.168.103.14"
    mLib$ = "RBPCSF"
    mNmtTar_DtfTxt = "tmpECL_DTFtxt"
    mIsByXls = False
    mSql$ = "Select LODTE AS yymd_EnterDte, ECL.* from ECL WHERE LPROD LIKE '07%' AND LQORD>LQSHP"
    mResult = jj.Bld_OdbcTwoQry_BySqlDtf(mSql, mIP, mLib, True, mNmtTar, mNmtTar_DtfTxt, mNmq1, mNmq2, mIsByXls, mFbTar, mRun)
Case 2
End Select
jj.Shw_Dbg cSub, cMod, "Result, mSql,mIP,mLib,mFbTar,mNmtTar,mNmtTar_DtfTxt,mNmq1,mNmq2,mIsByXls,mRun", mResult, mSql, mIP, mLib, mFbTar, mNmtTar, mNmtTar_DtfTxt, mNmq1, mNmq2, mIsByXls, mRun
Dim mDb As DAO.Database: If jj.Cv_Db_FmFb(mDb, mFbTar) Then Stop
Debug.Print jj.ToStr_LpAp(vbLf, "mNmq1.Sql,mNmq2.Sql", mDb.QueryDefs(mNmq1).Sql, mDb.QueryDefs(mNmq2).Sql)
If mFbTar <> "" Then jj.Cls_Db mDb
End Function
#End If
Function Bld_OdbcQs(pNmQs$ _
    , Optional pLm$ = "" _
    , Optional pLikNmDl$ = "*" _
    , Optional pFbTar$ = "" _
    , Optional pRun As Boolean = False _
    , Optional pDbQry As DAO.Database = Nothing _
    ) As Boolean
'Aim: Build and optionally not Exec a set of Odbc sql for {pNmqs$} with ref macro Rs from tmp{Nmqsns}_Prm & table tblOdbcSql.
'     Optional Keep or 2nd&3rd Qry Qry depend of jj.SysCfg_IsOdbcDb
''(0) [Run qryOdbc{pNmqs}_0*] to generate tmp{Nmqs}_Prm
''(1) [Build 3x query]  For each record in tblOdbcSql for {pNmqs$} build 3 queries (nn started from 11)
''                      1st Qry: {Nmqsns}_{nn}_0_{NmDl}      Select * from tmp{Nmqsns}_{NmDl}_Odbc
''                      2nd Qry: {Nmqsns}_{nn}_1_Fm_Below    Select * into tmp{Nmqsns}_{NmDl} from qry{Nmqsns}_{nn}_2_{NmDl}_Odbc
''                      3rd Qry: {Nmqsns}_{nn}_2_{NmDl}_Odbc (Fm tblOdbcSql, tmp{Nmqsns}_Prm, gIsLclMd
''                          Note: 2nd & 3rd may be multiple if there is multiple records in tmp{Nmqsns}_Prm.
''                                In this case
''                                2nd Qry: {Nmqsns}_{nn}_1_{i}Fm_Below    where {i}=0..9 for each rec in tmp{Nmqsns}_Prm
''                                3rd Qry: {Nmqsns}_{nn}_2_{i}{NmDl}_Odbc
''(2) Optional [Exec]
''(3) Optional [Rmv]
''Assume: In LclMd, the data table in the Odbc Sql should be linked by name without any pfx (eg. IIC, not RBPCSF_IIC)
'==Start
'(1) [Build 3x query]
Const cSub$ = "Bld_OdbcQs"

Dim mDbQry As DAO.Database: Set mDbQry = jj.Cv_Db(pDbQry)
'Test if tblOdbcSql exists and contain records for Nmqs='{pNmqs}'
Dim mCnt&: If jj.Fnd_RecCnt_ByNmtq(mCnt, "tblOdbcSql", jj.Q_S(pNmQs, "NmQs='*'"), pDbQry) Then ss.A 1: GoTo E
If mCnt = 0 Then Exit Function

'Build 3x Query
''Find xAnDl$(), xAySqlTp$(),xAySqlTp_LclMd$ from tblOdbcSql
Dim mSql$: mSql = "Select NmDl,SqlTp,SqlTp_LclMd from tblOdbcSql where Nmqs='" & pNmQs & cQSng
Dim xAnDl$(), xAySqlTp$(), xAySqlTp_LclMd$(): If jj.Fnd_LoAyV_FmSql_InDb(mDbQry, mSql, "NmDl,SqlTp,SqlTp_LclMd", xAnDl, xAySqlTp, xAySqlTp_LclMd) Then ss.A 2: GoTo E
Dim J%
If jj.SysCfg_IsLclMd Then
    For J = 0 To jj.Siz_Ay(xAySqlTp) - 1
        If xAySqlTp_LclMd(J) <> "" Then xAySqlTp(J) = xAySqlTp_LclMd(J)
    Next
End If

'Exec qryOdbc{Nmqsns}_0* to get {mNmtPrm} = tmp{NmQsns}_Prm
Dim mNmQsNs$: mNmQsNs = mID(pNmQs, 4)
Dim mNmtPrm$: mNmtPrm = "tmpOdbc" & mNmQsNs & "_Prm"
If jj.Crt_Tbl_tmpXXX_Prm_By_qryOdbcXXX_0(mNmQsNs, pLm) Then ss.A 3: GoTo E

'Find NPrm%, mRsPrm from {mNmtPrm}.
''Set mRsPrm & Get NPrm%
Dim mRsPrm As DAO.Recordset: Set mRsPrm = mDbQry.OpenRecordset("Select * from " & mNmtPrm)
Dim NPrm%: If jj.Fnd_RecCnt_ByRs(NPrm, mRsPrm) Then mRsPrm.Close: ss.A 4: GoTo E
If NPrm <= 0 Then ss.A 5, "Param table tmp{Nmqsns}_Prm has no record.   tmp{Nmqsns}_Prm is expected to be generated by RunQry qryOdbc{mNmqsns}_0*": GoTo E
 
'Find mAyDsn$(), mAyBoolIsByDtf(), mAyDtfIP$(), mAyDtfLib$(): 0 To NPrm - 1 from {mNmtPrm}
Dim mAyDsn$(), mAyIsByDtf$(), mAyDtfIP$(), mAyDtfLib$()
If jj.Fnd_LoAyV_FmRs(mRsPrm, "Dsn,IsByDtf,IP,Lib", mAyDsn, mAyIsByDtf, mAyDtfIP, mAyDtfLib) Then ss.A 7, "the table tmp{Nmqsns}_Prm must contain a field DSN.  tmp{Nmqsns}_Prm is expected to be generated by RunQry qryOdbc{mNmqsns}_0*": GoTo E
Dim N%: N = jj.Siz_Ay(mAyDsn)
ReDim mAyBoolIsByDtf(0 To N - 1) As Boolean
For J = 0 To N - 1
    mAyBoolIsByDtf(J) = CBool(mAyIsByDtf(J))
Next

'Dlt qryOdbc{mNmqsns}
If jj.Dlt_Qry_ByPfx("qryOdbc" & mNmQsNs & "_1") Then ss.A 8: GoTo E
If jj.Dlt_Qry_ByPfx("qryOdbc" & mNmQsNs & "_2") Then ss.A 9: GoTo E

''Build & Exec 3x Query by looping xAnDl(), xAySqlTp()
ReDim mAySql$(0 To NPrm - 1)
N = jj.Siz_Ay(xAySqlTp)

For J = 0 To N - 1
    '''Find mAySql() by mRsPrm, xAySqlTp(J)
    If xAnDl(J) Like pLikNmDl Then
        mRsPrm.MoveFirst
        Dim I%: I = 0
        While Not mRsPrm.EOF
            mAySql(I) = jj.Fmt_Str_ByRs(xAySqlTp(J), mRsPrm): I = I + 1
            mRsPrm.MoveNext
        Wend
        If I > 0 Then
            jj.Shw_Sts jj.Fmt_Str("{0}_{1} ({2} {3}) .. Site:{4}", pNmQs, xAnDl(J), J, N - 1, I)
        Else
            jj.Shw_Sts jj.Fmt_Str("{0}_{1} ({2} {3}) ..", pNmQs, xAnDl(J), J, N - 1)
        End If
        If jj.Bld_OdbcQs_ByAySelSql(pNmQs, J + 11, xAnDl(J), mAyDsn, mAyBoolIsByDtf, mAyDtfIP, mAyDtfLib, mAySql, pFbTar, pRun) Then ss.A 9: GoTo E
    End If
Next
jj.Clr_Sts
Exit Function
R: ss.R
E: Bld_OdbcQs = True: ss.B cSub, cMod, "pLm,pLikNmDl,pFbTar,pRun,pDbQry", pLm, pLikNmDl, pFbTar, pRun, jj.ToStr_Db(pDbQry)
End Function
#If Tst Then
Function Bld_OdbcQs_Tst() As Boolean
Const cSub$ = "Bld_OdbcQs_Tst"
Dim mCase As Byte: mCase = 2
Dim mResult As Boolean
Dim mPfxNmObj_Src$, mFb_Src$
Dim mNmQs$, mLm$, mFbTar$, mLikNmDl$, mRun As Boolean
mFbTar = "c:\aa.mdb"
Select Case mCase
Case 1  ' MPS
    mPfxNmObj_Src = "qryOdbcMPS"
    mFb_Src = "D:\SPLHalfWayHouse\MPSDetail\VerNew@2007_01_04\WorkingDir\PgmObj\MPS_Odbc.Mdb"
    If jj.Crt_Tbl_FmLnkNmt(mFb_Src, "tblOdbcSql,tblMPSPrm") Then ss.A 1: GoTo E
    If False Then If jj.Crt_Tbl_FmLnkNmt(jj.Sffn_This, "mstBrand,mstEnv,mstLib,mstIP") Then ss.A 2: GoTo E
    If jj.Cpy_Obj_ByPfx(mPfxNmObj_Src, acQuery, mFb_Src) Then ss.A 3: GoTo E
    
    mNmQs = "qryMPS"
    mLm = "Env=NAPROD,Brand=TH"
    mRun = False
Case 2  ' Fc
    mPfxNmObj_Src = "qryOdbcFc_0"
    mFb_Src = "D:\SPLHalfWayHouse\MPSDetail\VerNew@2007_01_04\WorkingDir\PgmObj\MPS_RfhFc.Mdb"
    If jj.Crt_Tbl_FmLnkLnt(mFb_Src, "tblOdbcSql,tblFcPrm") Then ss.A 1: GoTo E
    If False Then If jj.Crt_Tbl_FmLnkNmt(jj.Sffn_This, "mstBrand,mstEnv,mstLib,mstIP") Then ss.A 1: GoTo E
    If jj.Cpy_Obj_ByPfx(mPfxNmObj_Src, acQuery, mFb_Src) Then ss.A 1: GoTo E
    
    mNmQs = "qryFc"
    mLm = "Env=NAPROD,Brand=TH"
    mRun = False
End Select
mResult = jj.Bld_OdbcQs(mNmQs, mLm, , mFbTar, mRun)
jj.Shw_Dbg cSub, cMod, , "Result,mPfxNmObj_Src,mFb_Src", mResult, mPfxNmObj_Src, mFb_Src
Exit Function
R: ss.R
E: Bld_OdbcQs_Tst = True: ss.B cSub, cMod
End Function
#End If
Function Bld_OdbcQs_ByAySelSql(pNmQs$, pMajNo As Byte, pNmDl$, pAyDsn$(), _
        pAyIsByDtf() As Boolean, pAyDtfIP$(), pAyDtfLib$(), _
        pAySelSql$() _
        , Optional pFbOupTbl$ = "" _
        , Optional pRun As Boolean = False _
        ) As Boolean
Const cSub$ = "Bld_OdbcQs_ByAySelSql"
'Aim: Create 3 sets of queries (Q0,Q1,Q2) in {pDb}.  If {pRun}, table tmp{mNmqsns}_{pNmDl} in {pDb} will be created.
'Notes:
'     [mNmtTar] The all Q1 queries will download data to [mNmtTar] = tmp{mNmqsns}_{pNmDl} from multi-site as defined in pAy*.
'           # of queries in Q1 will be same as # of element in pAy*.
'           The first query of Q1 is Create Table query, while the rest are Append Table query
'     [ClearUp] If {pRun} & not {gOdbcDbg} then delete the [Q1 & Q2] & [Dtf import tables], so that only tmp{mNmqsns}_{pNmDl} will left
'     [DTF]     If the query is DTF, it always download data and import to the database.
'           [mNmtTar_DTF] = [tmp{mNmqsns}_{pNmDl}_DTFTXT{n}], will be import in CurrentDb
'                               n is 0 - N-1, N is # of element in pAy*
'           [Q2]: Select * from {[mNmtTar_DTF]}
'           [Q1]: Standard,   Const cSql_Crt$ = "Select {0} into {1} from {2}"
'                             Const cSql_App$ = "Insert into {1} Select {0} from {2}"
'     [Dsn]
'
'Assume: the pDsn is used to create the library, which means there will be no library in the ODBC query.
''    Name                                  Sql
''Q0: qryOdbc{pNmQs}_{pMajNo}_0_{pNmDl}     Select * from tmp{pNmQs}_{pNmDl}
''Q1: qryOdbc{pNmQs}_{pMajNo}_1_{n}Fm_Below Select * into tmp{pNmQs}_{pNmDl} from {qry3n}
''Q2: qryOdbc{pNmQs}_{pMajNo}_2_{n}{xNmDl}  {pAySelSql()}
If Left(pNmQs, 3) <> "qry" Then ss.A 1, "pNmqs must begins with qry": GoTo E
If Not pRun And Not jj.SysCfg_IsDbgOdbc Then ss.A 2, "Must either pRun or gIsOdbcDbg": GoTo E
Dim N%: N = jj.Siz_Ay(pAyIsByDtf)
If jj.Siz_Ay(pAyDtfIP) <> N Then ss.A 3, "Siz of pAyIsByDtf and pAyDtfIP must be the same", "SizOf pAyIsByDtf", "SizOf pAyDtfIP", N, jj.Siz_Ay(pAyDtfIP): GoTo E
If jj.Siz_Ay(pAyDtfLib) <> N Then ss.A 4, "Siz of pAyIsByDtf and pAyDtfLib must be the same", "SizOf pAyIsByDtf", "SizOf pAyDtfLib", N, jj.Siz_Ay(pAyDtfLib): GoTo E
If jj.Siz_Ay(pAySelSql) <> N Then ss.A 5, "Siz of pAyIsByDtf and pAySelSql must be the same", "SizOf pAyIsByDtf", "SizOf pAySelSql", N, jj.Siz_Ay(pAySelSql): GoTo E

Dim mMajNo$: mMajNo = Format(pMajNo, "00")
Dim mNmQsNs$: mNmQsNs = mID(pNmQs, 4)

'Build Query 0
Dim mNmq0$: mNmq0 = jj.Fmt_Str("qryOdbc{0}_{1}_0_{2}", mNmQsNs, mMajNo, pNmDl)
Dim mInFbOupTbl$: If pFbOupTbl <> "" Then mInFbOupTbl = jj.Q_S(pFbOupTbl, " in '*'")
Dim mSql$: mSql = jj.Fmt_Str("Select * from tmp{0}_{1}{2}", mNmQsNs, pNmDl, mInFbOupTbl)
If jj.Crt_Qry(mNmq0, mSql) Then ss.A 2: GoTo E
Dim J%: For J = 0 To N - 1
    Dim A$: A = IIf(N = 1, "", J)
     'Build Query 2
    Dim mNmq1$: mNmq1 = jj.Fmt_Str("qryOdbc{0}_{1}_1_{2}Fm_Below", mNmQsNs, mMajNo, A)
    Dim mNmq2$: mNmq2 = jj.Fmt_Str("qryOdbc{0}_{1}_2_{2}{3}", mNmQsNs, mMajNo, A, pNmDl)
    Dim mNmtTar$: mNmtTar = "tmp" & mNmQsNs & "_" & pNmDl
    Dim mNmtTar_DtfTxt$: mNmtTar_DtfTxt$ = "tmp" & mNmQsNs & "_" & pNmDl & "_" & J & "_DtfTxt"
    Dim mIsCrtTbl As Boolean: mIsCrtTbl = (J = 0)
    If pAyIsByDtf(J) Then
        If N = 1 Then
            If jj.Crt_Tbl_FmDTF_Sql(pAyDtfIP(J), pAySelSql(J), mNmtTar, pFbOupTbl, pAyDtfLib(J)) Then ss.A 4: GoTo E
        Else
            If jj.Bld_OdbcTwoQry_BySqlDtf(pAySelSql(J), pAyDtfIP(J), pAyDtfLib(J), mIsCrtTbl, mNmtTar, mNmtTar_DtfTxt, mNmq1, mNmq2, , pFbOupTbl, pRun) Then ss.A 4: GoTo E
        End If
    Else
        If jj.Bld_OdbcTwoQry_BySqlDsn(pAySelSql(J), pAyDsn(J), mIsCrtTbl, mNmtTar, pFbOupTbl, mNmq1, mNmq2, pRun) Then ss.A 4: GoTo E
    End If

    If pRun Then
        If Not jj.SysCfg_IsDbgOdbc Then
            jj.Dlt_Qry mNmq1
            jj.Dlt_Qry mNmq2
        End If
    End If
Next
Exit Function
R: ss.R
E: Bld_OdbcQs_ByAySelSql = True: ss.B cSub, cMod, "pNmQs,pMajNo,pNmDl,pAyDsn(),AyIsByDtf(),pAyDtfIP(),pAyDtfLib(),pAySelSql(),pFbOupTbl,pRun", pNmQs, pMajNo, pNmDl, jj.ToStr_Ays(pAyDsn()), jj.ToStr_AyBool(pAyIsByDtf()), jj.ToStr_Ays(pAyDtfIP()), jj.ToStr_Ays(pAyDtfLib()), jj.ToStr_Ays(pAySelSql()), pFbOupTbl, pRun
End Function
#If Tst Then
Function Bld_OdbcQs_ByAySelSql_Tst() As Boolean
Const cSub$ = "Bld_OdbcQs_ByAySelSql_Tst"
Dim N%: N = 1
Dim mNmQs$, mMajNo As Byte, mNmDl$, mFbOupTbl$
ReDim mAyDsn$(0 To N)
ReDim mAyIsByDtf(0 To N) As Boolean
ReDim mAyDtfIP$(0 To N)
ReDim mAyDtfLib$(0 To N)
ReDim mAySql$(0 To N)
Dim mCase As Byte: mCase = 1
Dim mResult As Boolean
Select Case mCase
Case 1
    mNmQs = "qryDD"
    mMajNo = 11
    mNmDl = "IIC"
    '
    mAyDsn(0) = "CHPROD_BPCSF"
    mAyIsByDtf(0) = True
    mAyDtfIP(0) = "192.168.103.26"
    mAyDtfLib(0) = "BPCSF"
    mAySql(0) = "Select 'CH' AS SRC, IIC.* from IIC"
    '
    mAyDsn(1) = "FEPROD_RBPCSF"
    mAyIsByDtf(1) = True
    mAyDtfIP(1) = "192.168.103.14"
    mAyDtfLib(1) = "RBPCSF"
    mAySql(1) = "Select 'FE' AS SRC, IIC.* from IIC"
Case 2
    'tmpDD_IIC will be created
    mNmQs = "qryDD"
    mMajNo = 11
    mNmDl = "IIC"
    '
    mAyDsn(1) = ""
    mAyIsByDtf(1) = True
    mAyDtfIP(1) = "192.168.103.13"
    mAyDtfLib(1) = "BPCSF"
    mAySql(1) = "Select 'US' AS SRC, IIC.* from IIC"
    '
    mAyDsn(0) = "FEPROD_RBPCSF"
    mAyIsByDtf(0) = False
    mAyDtfIP(0) = ""
    mAyDtfLib(0) = ""
    mAySql(0) = "Select 'FE' AS SRC, IIC.* from IIC"
Case 3
    mNmQs = "qryFc"
    mMajNo = 11
    mNmDl = "KMR"
    '
    mAyDsn(0) = ""
    mAyIsByDtf(0) = True
    mAyDtfIP(0) = "192.168.103.13"
    mAyDtfLib(0) = "BPCSF"
    mAySql(0) = "Select  'NA' AS Site, Trim(MPROD) As SKU, MRFAC As Fac, MRWHS As Whs," & _
        " MRDTE As yymd_FcBegDate," & _
        " MPDTE As yymd_FcEndDate," & _
        " MRCDT As yymd_EnterDate," & _
        " MQTY As FcQty" & _
        " From KMRL01" & _
        " Where MID='MR' AND MPROD LIKE '17%' AND MRDTE >= 20070201 AND MTYPE='F' AND MRFAC IN ('U1','C1')"
    '
    mAyDsn(1) = "FEPROD_RBPCSF"
    mAyIsByDtf(1) = False
    mAyDtfIP(1) = ""
    mAyDtfLib(1) = ""
    mAySql(1) = "Select 'FE' AS Site, Trim(MPROD) As SKU, MRFAC As Fac, MRWHS As Whs," & _
        " MRDTE As yymd_FcBegDate," & _
        " MPDTE As yymd_FcEndDate," & _
        " MRCDT As yymd_EnterDate," & _
        " MQTY As FcQty" & _
        " From KMRL01" & _
        " Where MID='MR' AND MPROD LIKE '17%' AND MRDTE >= 20070201 AND MTYPE='F' AND MRFAC IN ('R9','T9')"
End Select
mResult = jj.Bld_OdbcQs_ByAySelSql(mNmQs, mMajNo, mNmDl, mAyDsn, mAyIsByDtf, mAyDtfIP, mAyDtfLib, mAySql, mFbOupTbl)
jj.Shw_Dbg cSub, cMod, , "Result,mNmqs,mMajNo,mNmDl", mResult, mNmQs, mMajNo, mNmDl
End Function
#End If
Function Bld_OdbcQs_ByAySelSql_ByDsn(pNmQs$, pMajNo As Byte, pNmDl$, pAyDsn$(), pAySelSql$() _
        , Optional pFbOupTbl$ = "" _
        , Optional pRun As Boolean = False _
        ) As Boolean
Const cSub$ = "Bld_OdbcQs_ByAySelSql_ByDsn"
'Aim: Create 3 sets of queries (Q0,Q1,Q2) in CurrentDb.  If {pRun}, table tmp{mNmqsns}_{pNmDl} in {pFbOupTbl} will be created.
'Notes:
'     [mNmtTar] = tmp{mNmqsns}_{pNmDl} from multi-site as defined in pAy*.
'     Q1    # of queries in Q1 will be same as # of element in pAy*.
'           Running all Q1 will download data to [mNmtTar]
'           The first query of Q1 is Create Table query, while the rest are Append Table query
'     [ClearUp] If {pRun} & not {gOdbcDbg} then delete the [Q1 & Q2] & [Dtf import tables], so that only tmp{mNmqsns}_{pNmDl} will left
'           [Q2]: Select * from {[mNmtTar_DTF]}
'           [Q1]: Standard,   Const cSql_Crt$ = "Select {0} into {1} from {2}"
'                             Const cSql_App$ = "Insert into {1} Select {0} from {2}"
'
'Assume: the pDsn is used to create the library, which means there will be no library in the ODBC query.
''    Name                                  Sql
''Q0: qryOdbc{pNmQs}_{pMajNo}_0_{pNmDl}     Select * from tmp{pNmQs}_{pNmDl}
''Q1: qryOdbc{pNmQs}_{pMajNo}_1_{n}Fm_Below Select * into tmp{pNmQs}_{pNmDl} from {qry3n}
''Q2: qryOdbc{pNmQs}_{pMajNo}_2_{n}{xNmDl}  {pAySelSql()}
If Left(pNmQs, 3) <> "qry" Then ss.A 1, "pNmQs must begins with qry": GoTo E
If Not pRun And Not jj.SysCfg_IsDbgOdbc Then ss.A 1, "Must either pRun or gIsOdbcDbg": GoTo E
Dim N%: N = jj.Siz_Ay(pAyDsn)
If jj.Siz_Ay(pAySelSql) <> N Then ss.A 2, "Siz of pAyDsn and pAySelSql must be the same", , "SizOf pAyDsn, SizOf pAySelSql", N, jj.Siz_Ay(pAySelSql): GoTo E

Dim mMajNo$: mMajNo = Format(pMajNo, "00")
Dim mNmQsNs$: mNmQsNs = mID(pNmQs, 4)

'Build Query 0
Dim mNmq0$: mNmq0 = jj.Fmt_Str("qryOdbc{0}_{1}_0_{2}", mNmQsNs, mMajNo, pNmDl)
Dim mInFbOupTbl$: If pFbOupTbl <> "" Then mInFbOupTbl = jj.Q_S(pFbOupTbl, " in '*'")
Dim mSql$: mSql = jj.Fmt_Str("Select * from tmp{0}_{1}{2}", mNmQsNs, pNmDl, mInFbOupTbl)
If jj.Crt_Qry(mNmq0, mSql) Then ss.A 2: GoTo E

Dim J%: For J = 0 To N - 1
    Dim A$: A = IIf(N = 1, "", J)
     'Build Query 2
    Dim mNmq1$: mNmq1 = jj.Fmt_Str("qryOdbc{0}_{1}_1_{2}Fm_Below", mNmQsNs, mMajNo, A)
    Dim mNmq2$: mNmq2 = jj.Fmt_Str("qryOdbc{0}_{1}_2_{2}{3}", mNmQsNs, mMajNo, A, pNmDl)
    Dim mNmtTar$: mNmtTar = "tmp" & mNmQsNs & "_" & pNmDl
    Dim mNmtTar_DtfTxt$: mNmtTar_DtfTxt$ = "tmp" & mNmQsNs & "_" & pNmDl & "_" & J & "_DtfTxt"
    Dim mIsCrtTbl As Boolean: mIsCrtTbl = (J = 0)
    If jj.Bld_OdbcTwoQry_BySqlDsn(pAySelSql(J), pAyDsn(J), mIsCrtTbl, mNmtTar, pFbOupTbl, mNmq1, mNmq2, pRun) Then ss.A 4: GoTo E

    If pRun Then
        If Not jj.SysCfg_IsDbgOdbc Then
            jj.Dlt_Qry mNmq1
            jj.Dlt_Qry mNmq2
        End If
    End If
Next
Exit Function
R: ss.R
E: Bld_OdbcQs_ByAySelSql_ByDsn = True: ss.B cSub, cMod, "pNmQs,pMajNo,pNmDl", pNmQs, pMajNo, pNmDl
End Function
#If Tst Then
Function Bld_OdbcQs_ByAySelSql_ByDsn_Tst() As Boolean
Const cSub$ = "Bld_OdbcQs_ByAySelSql_ByDsn_Tst"
Dim N%: N = 1
Dim mNmQs$, mMajNo As Byte, mNmDl$, mFbOupTbl$
ReDim mAyDsn$(0 To N)
ReDim mAySql$(0 To N)
Dim mCase As Byte: mCase = 1
Dim mResult As Boolean
Select Case mCase
Case 1
    mNmQs = "qryDD"
    mMajNo = 11
    mNmDl = "IIC"
    '
    mAyDsn(0) = "CHPROD_BPCSF"
    mAySql(0) = "Select 'CH' AS SRC, IIC.* from IIC"
    '
    mAyDsn(1) = "FEPROD_RBPCSF"
    mAySql(1) = "Select 'FE' AS SRC, IIC.* from IIC"
Case 2
    'tmpDD_IIC will be created
    mNmQs = "qryDD"
    mMajNo = 11
    mNmDl = "IIC"
    '
    mAyDsn(1) = ""
    mAySql(1) = "Select 'US' AS SRC, IIC.* from IIC"
    '
    mAyDsn(0) = "FEPROD_RBPCSF"
    mAySql(0) = "Select 'FE' AS SRC, IIC.* from IIC"
Case 3
    mNmQs = "qryFc"
    mMajNo = 11
    mNmDl = "KMR"
    '
    mAyDsn(0) = ""
    mAySql(0) = "Select  'NA' AS Site, Trim(MPROD) As SKU, MRFAC As Fac, MRWHS As Whs," & _
        " MRDTE As yymd_FcBegDate," & _
        " MPDTE As yymd_FcEndDate," & _
        " MRCDT As yymd_EnterDate," & _
        " MQTY As FcQty" & _
        " From KMRL01" & _
        " Where MID='MR' AND MPROD LIKE '17%' AND MRDTE >= 20070201 AND MTYPE='F' AND MRFAC IN ('U1','C1')"
    '
    mAyDsn(1) = "FEPROD_RBPCSF"
    mAySql(1) = "Select 'FE' AS Site, Trim(MPROD) As SKU, MRFAC As Fac, MRWHS As Whs," & _
        " MRDTE As yymd_FcBegDate," & _
        " MPDTE As yymd_FcEndDate," & _
        " MRCDT As yymd_EnterDate," & _
        " MQTY As FcQty" & _
        " From KMRL01" & _
        " Where MID='MR' AND MPROD LIKE '17%' AND MRDTE >= 20070201 AND MTYPE='F' AND MRFAC IN ('R9','T9')"
End Select
mResult = jj.Bld_OdbcQs_ByAySelSql_ByDsn(mNmQs, mMajNo, mNmDl, mAyDsn, mAySql, mFbOupTbl)
jj.Shw_Dbg cSub, cMod, "Result,mNmqs,mMajNo,mNmDl,mAyDsn,mAySql", mResult, mNmQs, mMajNo, mNmDl, jj.ToStr_Ays(mAyDsn), jj.ToStr_Ays(mAySql, , vbLf)
End Function
#End If
Function Bld_Sql_Sel_CurHostDta(oSql$, pRsUlSrc As DAO.Recordset, pNmtHost$, Optional oLExpr$) As Boolean
'Aim: Build {oSql} to get data from Host by referring {pRsUlSrc}
'{pRsUlSrc} fmt: First N field is PK, Then [Changed], Then pair [xxx], [New xxx]
Const cSub$ = "Bld_Sql_Sel_CurHostDta"
Dim mSel$
oLExpr = ""
Dim J%: For J = 0 To pRsUlSrc.Fields.Count - 1
    If pRsUlSrc.Fields(J).Name = "Changed" Then
        Dim I%: For I = J + 1 To pRsUlSrc.Fields.Count - 1 - 5 Step 2 'Skip 5 columns at end
            mSel = jj.Add_Str(mSel, pRsUlSrc.Fields(I).Name)
        Next
        Exit For
    End If
    With pRsUlSrc.Fields(J)
        Dim mA$: If jj.Join_NmV(mA, .Name, .Value) Then ss.A 1: GoTo E
    End With
    oLExpr = jj.Add_Str(oLExpr, mA, " and ")
Nxt:
Next
If mSel = "" Then ss.A 1, "mSel should be blank": GoTo E
If oLExpr = "" Then ss.A 2, "oLExpr should be blank": GoTo E
oSql = jj.Fmt_Str("Select {0} from {1} where {2}", mSel, pNmtHost, oLExpr)
Exit Function
R: ss.R
E: Bld_Sql_Sel_CurHostDta = True: ss.B cSub, cMod, "pRsUlSrc,pNmtHost", jj.ToStr_Flds(pRsUlSrc.Fields), pNmtHost
End Function
Function Bld_Sql_Upd_ByRs(oSqlUpd$, pRs As DAO.Recordset, pNmtTar$, pLmPk$, Optional pLnFld$ = "") As Boolean
''Aim: Build a {oSqlUpd} to Update table {pNmtTar} by the context in current record in {pRs}.
'      If {pLnFld} is given, only those fields in the list will be Updated.
'      If {pLnFld} is not given, all fields in {pRs} will be Updated
Const cSub$ = "Bld_Sql_Upd_ByRs"
Dim mLnFld$: If jj.Substract_Lst(mLnFld, jj.ToStr_Flds(pRs.Fields), pLmPk) Then ss.A 1: GoTo E
Dim mSet$: If jj.Set_Lv_ByRs(mSet, pRs, mLnFld) Then ss.A 1: GoTo E
Dim mCndn$: If jj.Set_Lv_ByRs(mCndn, pRs, pLmPk$, , " and ") Then ss.A 1: GoTo E
oSqlUpd = jj.ToSql_Upd(pNmtTar, mSet, mCndn)
Exit Function
R: ss.R
E: Bld_Sql_Upd_ByRs = True: ss.B cSub, cMod, "pRs,pNmtTar,pLmPk,pLnFld", pRs, pNmtTar, pLmPk, pLnFld
End Function
Function Bld_Sql_Upd_ByRsUlSrc(oSqlUpd$, pRsUlSrc As DAO.Recordset, pNmtHost$) As Boolean
'Aim: Build {oSql} to get data from Host by referring {pRsUlSrc}
'{pRsUlSrc} fmt: First N field is PK, Then [Changed], Then pair [xxx], [New xxx]
Const cSub$ = "Bld_Sql_Upd_ByRsUlSrc"
Dim mSet$, mCndn$, mA$
Dim J%: For J = 0 To Fct.MinInt(10, pRsUlSrc.Fields.Count - 1)
    If pRsUlSrc.Fields(J).Name = "Changed" Then
        Dim I%: For I = J + 2 To pRsUlSrc.Fields.Count - 1 - 5 Step 2 'Skip 5 columns at end
            Dim mFld As DAO.Field: Set mFld = pRsUlSrc.Fields(I)
            If Not IsNull(mFld.Value) Then
                If Left(mFld.Name, 4) <> "New " Then ss.A 1, "The I-th field is not beging [New ]", , "I,NmFld", I, mFld.Name: GoTo E
                If jj.Join_NmV(mA, mID(mFld.Name, 5), mFld.Value) Then ss.A 2, "The I-th field cannot build 'Set xx=xx'", , "I,NmFld", I, mFld.Name: GoTo E
                mSet = jj.Add_Str(mSet, mA)
            End If
        Next
        Exit For
    End If
    With pRsUlSrc.Fields(J)
        If jj.Join_NmV(mA, .Name, .Value) Then ss.A 1: GoTo E
    End With
    mCndn = jj.Add_Str(mCndn, mA, " and ")
Nxt:
Next
If mSet = "" Then ss.A 3, "mSet should be blank": GoTo E
If mCndn = "" Then ss.A 4, "mCndn should be blank": GoTo E
oSqlUpd = jj.Fmt_Str("Update {0} set {1} where {2}", pNmtHost, mSet, mCndn)
Exit Function
R: ss.R
E: Bld_Sql_Upd_ByRsUlSrc = True: ss.B cSub, cMod, "pRsUlSrc", jj.ToStr_Flds(pRsUlSrc.Fields)
End Function
Function Bld_Sql_Upd_ByRsUlSrc_Tst() As Boolean
Const cSub$ = "Bld_Sql_Upd_ByRsUlSrc_Tst"
Dim mRs As DAO.Recordset, mSql$
Dim mRslt As Boolean, mCase As Byte: mCase = 1
Select Case mCase
Case 1
    If False Then
        If jj.Crt_Tbl_ForEdtTbl("tblUsr", 1) Then ss.A 1: GoTo E
    End If
    Set mRs = CurrentDb.OpenRecordset("Select * from tmpEdt_tblUsr")
    With mRs
        While Not .EOF
            mRslt = jj.Bld_Sql_Upd_ByRsUlSrc(mSql, mRs, "tmpEdt_tblUsr") ', mAmFm, mAmTo)
            jj.Shw_Dbg cSub, cMod, , "mRslt,mSql", mRslt, mSql
            .MoveNext
        Wend
        .Close
    End With
Case 2
End Select
Exit Function
R: ss.R
E: Bld_Sql_Upd_ByRsUlSrc_Tst = True: ss.B cSub, cMod
End Function
Function Bld_Tbl_NRec2Lst(pNmtFm$, pNmtTo$, pLoKey$, pNmFld_NRec$, Optional pSepChr$ = cCommaSpc) As Boolean
'Aim: Create {pNmtTo} from {pNmtFm}.  The new table will have Key fields as list in {pLoKey} plus 1 more field [Lst{pNmFld}].
'     The value of this field [Lst{pNmFld}] is coming those records in {pNmtFm} of the current key.
Const cSub$ = "Bld_Tbl_NRec2Lst"
'Build Empty {pNmtTo}
If jj.Dlt_Tbl(pNmtTo) Then ss.A 1: GoTo E
Dim mSql$: mSql = jj.Fmt_Str("Select Distinct {0} into {1} from {2} where False", pLoKey, pNmtTo, pNmtFm)
If jj.Run_Sql(mSql) Then ss.A 1: GoTo E
mSql = jj.Fmt_Str("Alter table {0} Add COLUMN Lst{1} Memo", pNmtTo, pNmFld_NRec)
If jj.Run_Sql(mSql) Then ss.A 2: GoTo E
'Loop {pNmtFmt} having break @ mAnFldKey()
Dim mAnFldKey$(): mAnFldKey = Split(pLoKey, cComma)
Dim NKey%: NKey = jj.Siz_Ay(mAnFldKey)
Dim mRs As DAO.Recordset: Set mRs = CurrentDb.OpenRecordset(jj.Fmt_Str("Select {0},{1} from {2} order by {0},{1}", pLoKey, pNmFld_NRec, pNmtFm))
With mRs
    If .AbsolutePosition = -1 Then .Close: Exit Function
    ReDim mAyLasKeyVal(NKey - 1)
    Dim mLasKeyVal$:
    Dim J%: For J = 0 To NKey - 1
        mAyLasKeyVal(J) = .Fields(mAnFldKey(J)).Value
    Next
    If jj.Fnd_LvFmRs(mLasKeyVal, mRs, pLoKey) Then ss.A 3: GoTo E
    
    Dim mLst$
    Dim mSql_Tp$: mSql_Tp = jj.Fmt_Str("Insert into {0} ({1},Lst{2}) values ", pNmtTo, pLoKey, pNmFld_NRec) & "({0},'{1}')"
    While Not .EOF
        'If !Dte = #2/22/2007# And !Txt = "1010" And !InstId = 10 Then Stop
        If jj.IsSamKey_ByAnFldKey(mRs, mAnFldKey, mAyLasKeyVal) Then
            mLst = jj.Add_Str(mLst, .Fields(pNmFld_NRec).Value, pSepChr)
        Else
            '
            mSql = jj.Fmt_Str(mSql_Tp, mLasKeyVal, mLst)
            If jj.Run_Sql(mSql) Then ss.A 4: GoTo E

            For J = 0 To NKey - 1
                mAyLasKeyVal(J) = .Fields(mAnFldKey(J)).Value
            Next
            If jj.Fnd_LvFmRs(mLasKeyVal, mRs, pLoKey) Then ss.A 3: GoTo E
            mLst = mRs.Fields(pNmFld_NRec).Value
        End If
        .MoveNext
    Wend
    mSql = jj.Fmt_Str(mSql_Tp, mLasKeyVal, mLst)
    If jj.Run_Sql(mSql) Then ss.A 5: GoTo E
    .Close
End With
Exit Function
R: ss.R
E: Bld_Tbl_NRec2Lst = True: ss.B cSub, cMod, "pNmtFm,pNmtTo,pLoKey,pNmFld_NRec,pSepChr", pNmtFm, pNmtTo, pLoKey, pNmFld_NRec, pSepChr
End Function
Function Bld_Tbl_NRec2Lst_Tst() As Boolean
Const cSub$ = "Bld_Tbl_NRec2Lst_Tst"
jj.Shw_Dbg cSub, cMod, "Create {pNmtTo} from {pNmtFm}.  The new table will have Key fields list {pLoKey} plus 1 more field [Lst{pNmFld}], which has from NRec of {pNmtFm}"
Dim mNmtFm$: mNmtFm = "tmpBldTbl_NRec2Lst_Fm"
Dim mNmtTo$: mNmtTo = "tmpBldTbl_NRec2Lst_To"
Dim mLoKey$: mLoKey = "InstId,Dte,Txt"
Dim mNmFld_NRec$: mNmFld_NRec = "Num"
Dim mSepChr$: mSepChr = cComma

Dim mBldTstTbl As Boolean: mBldTstTbl = False
If mBldTstTbl Then
    If jj.Dlt_Tbl(mNmtFm) Then ss.A 1: GoTo E
    jj.Shw_Dbg cSub, cMod, "Create {pNmtTo} from {pNmtFm}.  The new table will have Key fields list {pLoKey} plus 1 more field [Lst{pNmFld}], which has from NRec of {pNmtFm}", _
        "mNmtFm,mNmtFm,mLoKey,mNmFld_NRec,mSepChr", mNmtFm, mNmtFm, mLoKey, mNmFld_NRec, mSepChr
    If jj.Run_Sql(jj.Fmt_Str("Create table {0} (InstId Long, Dte Date, Txt Text(10), Num Long)", mNmtFm)) Then ss.A 1: GoTo E
    Dim iInstId%: For iInstId = 0 To 10
        Dim iDte%: For iDte = 0 To 10
            Dim iTxt%: For iTxt = 1000 To 1010
                Dim iNum%: For iNum = 2000 To 2010
                    If jj.Run_Sql(jj.Fmt_Str("insert into {0} (InstId, Dte, Txt, Num) values ({1}, #{2}#, '{3}', {4})", mNmtFm, iInstId, Date + iDte, iTxt, iNum)) Then ss.A 1: GoTo E
                Next
            Next
        Next
    Next
End If
If jj.Bld_Tbl_NRec2Lst(mNmtFm, mNmtTo, mLoKey, mNmFld_NRec, mSepChr) Then ss.A 3: GoTo E
Exit Function
R: ss.R
E: Bld_Tbl_NRec2Lst_Tst = True: ss.B cSub, cMod
End Function
Function Bld_TblDte() As Boolean
Const cSub$ = "Bld_TblDte"
If jj.Dlt_Tbl("tblDte") Then ss.A 1: GoTo E
If jj.Run_Sql("Create table tblDte (Dte Date, YY Byte, MM Byte, DD Byte, [Wk#] Byte, [Wk Day] Text(3))") Then ss.A 2: GoTo E
Dim mRs As DAO.Recordset:
Set mRs = CurrentDb.TableDefs("tblDte").OpenRecordset
With mRs
    Dim J%: For J = 0 To 10000
        .AddNew
        !Dte = #1/1/2006# + J
        !yy = Year(!Dte) - 2000
        !MM = Month(!Dte)
        !DD = Day(!Dte)
        .Fields("Wk#").Value = Format(!Dte, "ww")
        .Fields("Wk Day").Value = Format(!Dte, "ddd")
        .Update
    Next
    .Close
End With
Exit Function
R: ss.R
E: Bld_TblDte = True: ss.B cSub, cMod
End Function
Private Function Bld_zFndDtfTp(oDtfTp$, pIP$, pLib$, pSql$, pConvTp$, pFfnFdf, pFfnTar, pPCFilTyp$, pSavFDF$) As Boolean
Const cSub$ = "Bld_zFndDtfTp"
If jj.Fnd_ResStr(oDtfTp, "DtfTp", True) Then ss.A 1: GoTo E
oDtfTp = jj.Fmt_Str_ByLpAp(oDtfTp, "IP,Lib,Sql,ConvTp,FfnFDF,FfnTar,PCFilTyp,SavFDF", pIP$, pLib$, Replace(Replace(pSql$, vbLf, " "), vbCr, " "), pConvTp$, pFfnFdf, pFfnTar, pPCFilTyp$, pSavFDF$)
Exit Function
R: ss.R
E: Bld_zFndDtfTp = True: ss.B cSub, cMod, "pIP,pLib,pSql,pConvTp,pFfnFdf,pFfnTar,pPCFilTyp,pSavFDF", pIP, pLib, pSql, pConvTp, pFfnFdf, pFfnTar, pPCFilTyp, pSavFDF
End Function


