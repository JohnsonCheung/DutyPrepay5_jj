Attribute VB_Name = "xToStr"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xToStr"
Function ToStr_App$(pApp As Application)
On Error GoTo R
If TypeName(pApp) = "Nothing" Then ToStr_App = "Nothing": Exit Function
ToStr_App = pApp.Name
Exit Function
R: ToStr_App = "Err: jj.ToStr_Acs(pApp). Msg=" & Err.Description
End Function
'Function ToStr_Doc$(pDoc As MSXML2.DOMDocument60)
'Const cSub$ = "ToStr_Doc"
'On Error GoTo R
'ToStr_Doc = pDoc.XML
'R: ToStr_Doc = "Err: jj.ToStr_Doc(pDoc). Msg=" & Err.Description
'End Function
Function ToStr_PgmADcl$(pArgDcl$ _
    , Optional pIsOpt As Boolean = False, Optional pIsByVal As Boolean = False, Optional pDftVal$ = "")
Dim mA$: mA = IIf(pIsByVal, "ByVal ", "") & pArgDcl & jj.Cv_Str(pDftVal, "=")
If pIsOpt Then ToStr_PgmADcl = jj.Q_S(mA, "[]"): Exit Function
ToStr_PgmADcl = mA
End Function
Function ToStr_Sgnt$(pNmTypArg$ _
    , Optional pIsOpt As Boolean = False _
    , Optional pIsByVal As Boolean = False _
    , Optional pIsAy As Boolean = False _
    , Optional pIsPrmAy As Boolean _
    , Optional pDftVal$ = "" _
    )
Dim mA$: mA = jj.Cv_Bool(pIsPrmAy, "PrmAy ") & pNmTypArg & jj.Cv_Bool(pIsAy, "()") & jj.Cv_Str(pDftVal, "=")
If pIsOpt Then ToStr_Sgnt = jj.Q_S(mA, "[]"): Exit Function
ToStr_Sgnt = mA
End Function
#If Tst Then
Function ToStr_Sgnt_Tst() As Boolean
Debug.Print ToStr_Sgnt("sdlkfj", True, True, True, True, "skdjf")
Debug.Print ToStr_Sgnt("sdlkfj")
End Function
#End If
Function ToStr_ArgDcl$(pNmArg$ _
    , Optional pNmTypArg$ = "$", Optional pIsPrmAy As Boolean = False, Optional pIsAy As Boolean = False)
Select Case pNmTypArg$
Case "$", "%", "!", "#", "&": ToStr_ArgDcl = pNmArg & pNmTypArg & IIf(pIsAy, "()", "")
Case Else:                    ToStr_ArgDcl = pNmArg & IIf(pIsAy, "()", "") & ":" & pNmTypArg
End Select
End Function
Function ToStr_Xls$(pXls As Excel.Application)
On Error GoTo R
ToStr_Xls = "(" & pXls.Workbooks.Count & ") Wb. Wb1=" & jj.ToStr_Wb(pXls.Workbooks(1))
Exit Function
R: ToStr_Xls = "Err: jj.ToStr_Xls(pXls). Msg=" & Err.Description
End Function
Function ToStr_TypFct$(pTypFct As eTypFct)
Select Case pTypFct
Case eTypFct.eFct: ToStr_TypFct = "Function "
Case eTypFct.eSub: ToStr_TypFct = "Sub "
Case eTypFct.eGet: ToStr_TypFct = "Property Get "
Case eTypFct.eLet: ToStr_TypFct = "Property Let "
Case eTypFct.eSet: ToStr_TypFct = "Property Set "
Case Else: ToStr_TypFct = "Unknown TypFct(" & pTypFct & ")"
End Select
End Function
Function ToStr_DPgmPrm$(pDPgm As d_Pgm, pAyDPrm() As d_Arg)
If jj.IsNothing(pDPgm) Then ToStr_DPgmPrm = "--Nothing--"
Dim mA$
With pDPgm
    Dim mNmRetTyp$, mNmAs
    Select Case .x_NmTypRet
    Case "#", "$", "%", "!", "&": mNmRetTyp = .x_NmTypRet
    Case Else: mNmAs = " As " & .x_NmTypRet
    End Select
    ToStr_DPgmPrm = jj.Q_MrkUp(jj.Fmt_Str("{0}{1}{2}{3}(){4}" _
        , IIf(.x_IsPrivate, "Private ", "") _
        , jj.ToStr_TypFct(.x_TypFct) & " " _
        , jj.Join_Lv(".", .x_NmPrj, .x_Nmm, .x_NmPrc) _
        , mNmRetTyp _
        , mNmAs), "DPgm") & vbLf & jj.Q_MrkUp(.x_Aim, "Aim")
End With
End Function
Function ToStr_DPgm$(pDPgm As d_Pgm, Optional pByEle As Boolean = False)
If jj.IsNothing(pDPgm) Then ToStr_DPgm = "--Nothing--"
Dim mA$
With pDPgm
    Dim mNmTypRet$, mNmAs
    Select Case .x_NmTypRet
    Case "#", "$", "%", "!", "&": mNmTypRet = .x_NmTypRet
    Case Else: mNmAs = " As " & .x_NmTypRet
    End Select
    ToStr_DPgm = jj.Q_MrkUp(jj.Fmt_Str("{0}{1}{2}{3}(){4}" _
        , IIf(.x_IsPrivate, "Private ", "") _
        , jj.ToStr_TypFct(.x_TypFct) & " " _
        , jj.Join_Lv(".", .x_NmPrj, .x_Nmm, .x_NmPrc) _
        , mNmTypRet _
        , mNmAs), "DPgm") & vbLf & jj.Q_MrkUp(.x_Aim, "Aim")
End With
End Function
Function ToStr_DArg$(pDArg As d_Arg, Optional pByEle As Boolean = False)
If TypeName(pDArg) = "Nothing" Then ToStr_DArg = "Nothing": Exit Function
Dim mA$
With pDArg
    If pByEle Then
        ToStr_DArg = jj.Fmt_Str("IsOpt={0}|IsByVal={1}|NmArg={2}|NmTypArg={3}|DftVal={4}" _
            , .x_IsOpt, .x_IsByVal, .x_NmArg, .x_NmTypArg, .x_DftVal)
        Exit Function
    End If
    If .x_IsOpt Then mA = "Optional "
    If .x_IsByVal Then mA = mA & ""
    mA = mA & .x_NmArg
    Select Case .x_NmTypArg
    Case "#", "$", "%", "&", "!":
        If Right(.x_NmArg, 2) = "()" Then
            mA = Left(mA, Len(mA) - 2) & .x_NmTypArg & "()"
        Else
            mA = mA & .x_NmTypArg
        End If
    Case "String"
        If Right(.x_NmArg, 2) = "()" Then
            mA = Left(mA, Len(mA) - 2) & "$()"
        Else
            mA = mA & "$"
        End If
    Case "Integer"
        If Right(.x_NmArg, 2) = "()" Then
            mA = Left(mA, Len(mA) - 2) & "%()"
        Else
            mA = mA & "%"
        End If
    Case "Single"
        If Right(.x_NmArg, 2) = "()" Then
            mA = Left(mA, Len(mA) - 2) & "!()"
        Else
            mA = mA & "!"
        End If
    Case "Long"
        If Right(.x_NmArg, 2) = "()" Then
            mA = Left(mA, Len(mA) - 2) & "&()"
        Else
            mA = mA & "&"
        End If
    Case "Double"
        If Right(.x_NmArg, 2) = "()" Then
            mA = Left(mA, Len(mA) - 2) & "#()"
        Else
            mA = mA & "#"
        End If
    Case Else: mA = mA & " As " & .x_NmTypArg
    End Select
    Select Case VarType(.x_DftVal)
    Case vbEmpty, vbNull
    Case vbString:
        If Trim(.x_DftVal) <> "" Then mA = mA & " = " & .x_DftVal
    Case Else: Stop
    End Select
End With
ToStr_DArg = mA
End Function
Function ToStr_AyDArg$(pAyDArg() As d_Arg, Optional pByEle As Boolean = False)
Dim N%: N = jj.Siz_AyDArg(pAyDArg)
If N = 0 Then ToStr_AyDArg = "--NoArg--": Exit Function
Dim J%, mA$
For J = 0 To N - 1
    mA = jj.Add_Str(mA, jj.ToStr_DArg(pAyDArg(J), pByEle), vbLf)
Next
ToStr_AyDArg = mA
End Function
Function ToStr_Md$(pMd As CodeModule)
On Error GoTo R
ToStr_Md = pMd.Parent.Collection.Parent.Name & "." & pMd.Name
Exit Function
R: ToStr_Md = "Err: jj.ToStr_Md(pMd). Msg=" & Err.Description
End Function
Function ToStr_Prj$(pPrj As VBProject)
On Error GoTo R
ToStr_Prj = pPrj.Name
Exit Function
R: ToStr_Prj = "Err: jj.ToStr_Prj(pPrj). Msg=" & Err.Description
End Function
Function ToStr_TypCmp$(pTypCmp As VBIDE.vbext_ComponentType)
Select Case pTypCmp
Case VBIDE.vbext_ComponentType.vbext_ct_ActiveXDesigner:    ToStr_TypCmp = "ActX"
Case VBIDE.vbext_ComponentType.vbext_ct_ClassModule:        ToStr_TypCmp = "Class"
Case VBIDE.vbext_ComponentType.vbext_ct_Document:           ToStr_TypCmp = "Doc"
Case VBIDE.vbext_ComponentType.vbext_ct_MSForm:             ToStr_TypCmp = "Frm"
Case VBIDE.vbext_ComponentType.vbext_ct_StdModule:          ToStr_TypCmp = "Mod"
Case Else: ToStr_TypCmp = "Unknow(" & pTypCmp & ")"
End Select
End Function
Function ToStr_Vayv$(pVayv, Optional pQ$ = "", Optional pSepChr$ = cCommaSpc)
If TypeName(pVayv) = "Nothing" Then ToStr_Vayv = "": Exit Function
Dim mAy(): mAy = pVayv
ToStr_Vayv = jj.ToStr_AyV(mAy, pQ, pSepChr)
End Function
Function ToStr_SqlLnt$(pSql$)
Const cSub$ = "ToStr_SqlLnt"
On Error GoTo R
Dim mAnt$()
If jj.Brk_Sql_ToAnt(mAnt, pSql) Then ss.A 1: GoTo E
ToStr_SqlLnt = Join(mAnt, cComma)
Exit Function
R: ss.R
E: ss.B cSub, cMod, "pSql", pSql
End Function
Function ToStr_Acs$(pAcs As Access.Application)
Const cSub$ = "ToStr_Acs"
On Error GoTo R
If TypeName(pAcs) = "Nothing" Then ToStr_Acs = "Nothing": Exit Function
ToStr_Acs = pAcs.CurrentDb.Name
Exit Function
R: ss.R
E: ss.B cSub, cMod
End Function
Function ToStr_Pkey$(pNmt$)
Const cSub$ = "ToStr_Pkey"
'Aim: Find the Pkey of given pNmt
On Error GoTo R
Dim I%
For I% = 0 To CurrentDb.TableDefs(pNmt).Indexes.Count - 1
    If CurrentDb.TableDefs(pNmt).Indexes(I).Primary Then
        Dim J%, mA$
        For J = 0 To CurrentDb.TableDefs(pNmt).Indexes(I).Fields.Count - 1
            mA = jj.Add_Str(mA, CurrentDb.TableDefs(pNmt).Indexes(I).Fields(J).Name)
        Next
        ToStr_Pkey = mA
        Exit Function
    End If
Next
ss.A 1, "No Pkey index"
R: ss.R
E: ss.B cSub, cMod, "pNmt", pNmt
End Function
#If Tst Then
Function ToStr_Pkey_Tst() As Boolean
Debug.Print jj.ToStr_Pkey("mstAllBrand")
Debug.Print jj.ToStr_Pkey("mstAllBrandaa")
End Function
#End If
Function ToStr_NmV$(pNm$, pV)
ToStr_NmV = pNm & "=[" & pV & "]"
End Function
Function ToStr_HostSts$(pHostSts As eHostSts)
Dim mA$
Select Case pHostSts
Case e1Rec: mA = "e1Rec"
Case e0Rec: mA = "e0Rec"
Case e2Rec: mA = "e2Rec"
Case eHostCpyToFrm: mA = "eHostCpyToFrm"
Case eUnExpectedErr: mA = "eUnExpectedErr"
Case Else: mA = "jj.ToStr_HostSts: " & pHostSts
End Select
ToStr_HostSts = mA
End Function
Function ToStr_AyLng$(pAyLng&())
Dim J%, N%: N = jj.Siz_Ay(pAyLng): If N = 0 Then Exit Function
Dim mS$: mS = pAyLng(0)
For J = 1 To N - 1
    mS = mS & ", " & pAyLng(J)
Next
ToStr_AyLng = mS
End Function
Function ToStr_An2V_New$(pAn2V() As tNm2V, Optional pSkipNull As Boolean = False)
Dim J%, mA$, V
For J = 0 To jj.Siz_An2V(pAn2V) - 1
    V = pAn2V(J).NewV
    If pSkipNull Then If IsNull(V) Then GoTo Nxt
    mA = jj.Add_Str(mA, jj.Q_V(V))
Nxt:
Next
ToStr_An2V_New = mA
End Function
Function ToStr_An2V$(pAn2V() As tNm2V, Optional pSepChr$ = vbLf)
Dim J%, mA$
For J = 0 To jj.Siz_An2V(pAn2V) - 1
    mA = jj.Add_Str(mA, jj.ToStr_Nm2V(pAn2V(J)), pSepChr)
Next
ToStr_An2V = mA
End Function
Function ToStr_Nm2V_Set(pNm2V As tNm2V) As Boolean
With pNm2V
    ToStr_Nm2V_Set = .Nm & "=" & jj.Q_V(.NewV)
End With
End Function
Function ToStr_Nm2V$(pNm2V As tNm2V)
Dim mIsEq As Boolean
With pNm2V
    If jj.IfEq_Nm2V(mIsEq, pNm2V) Then GoTo E
    Dim mA$
    If mIsEq Then
        mA$ = "<-NoChg"
    Else
        mA$ = "<-[" & jj.Q_V(.OldV) & "]"
    End If
    ToStr_Nm2V = .Nm & "=[" & jj.Q_V(.NewV) & "]" & mA
End With
Exit Function
E: ToStr_Nm2V = "Er IsEq_Nm2V"
End Function
Function ToStr_Am$(pAm() As tMap _
    , Optional pBrkChr$ = "=" _
    , Optional pQ1$ = "" _
    , Optional pQ2$ = "" _
    , Optional pSepChr$ = cCommaSpc _
    )
Dim N%: N% = jj.Siz_Am(pAm): If N% = 0 Then Exit Function
Dim J%, A$
For J = 0 To N - 1
    With pAm(J)
        If .F1 = "" Then
            If .F2 = "" Then
                A = jj.Add_Str(A, "", pSepChr)
            Else
                A = jj.Add_Str(A, jj.Q_S(.F2, pQ2), pSepChr)
            End If
        Else
            If .F2 = "" Then
                A = jj.Add_Str(A, jj.Q_S(.F1, pQ1), pSepChr)
            Else
                A = jj.Add_Str(A, jj.Q_S(.F1, pQ1) & pBrkChr & jj.Q_S(.F2, pQ2), pSepChr)
            End If
        End If
    End With
Next
ToStr_Am = A
End Function
Function ToStr_AmF1$(pAm() As tMap, Optional pQ$ = "", Optional pSepChr$ = cCommaSpc)
Dim N%: N = jj.Siz_Am(pAm): If N% = 0 Then Exit Function
Dim J%: For J = 0 To N - 1
    Dim A$: A = jj.Add_Str(A, jj.Q_S(pAm(J).F1, pQ), pSepChr)
Next
ToStr_AmF1 = A
End Function
#If Tst Then
Function ToStr_AmF1_Tst() As Boolean
Const cSub$ = "ToStr_AmF1_Tst"
Const cLm$ = "aaa=xxx,bbb=yyy,1111"
Dim mAm() As tMap: mAm = Get_Am_ByLm(cLm)
jj.Shw_Dbg cSub, cMod
Debug.Print "Input-----"
Debug.Print "cLm: "; cLm$
Debug.Print "Output----"
Debug.Print "jj.ToStr_AmF1(cLm): "; jj.ToStr_AmF1(mAm)
End Function
#End If
Function ToStr_AmF2$(pAm() As tMap, Optional pQ$ = "", Optional pSepChr$ = cCommaSpc)
'Aim: list the F2 of {pAm} ToStr
Dim N%: N = jj.Siz_Am(pAm): If N% = 0 Then Exit Function
Dim J%: For J = 0 To N - 1
    Dim A$: A = jj.Add_Str(A, jj.Q_S(pAm(J).F2, pQ), pSepChr)
Next
ToStr_AmF2 = A
End Function
#If Tst Then
Function ToStr_AmF2_Tst() As Boolean
Const cSub$ = "ToStr_AmF2_Tst"
Const cLm$ = "aaa=xxx,bbb=yyy,1111"
Dim mAm() As tMap: mAm = jj.Get_Am_ByLm(cLm)
jj.Shw_Dbg cSub, cMod
Debug.Print "Input-----"
Debug.Print "cLm: "; cLm$
Debug.Print "Output----"
Debug.Print "jj.ToStr_AmF2(cLm): "; jj.ToStr_AmF2(mAm)
End Function
#End If
Function ToStr_AyNm2V$(pAyNm2V() As tNm2V)
Stop
End Function
Function ToStr_Ays$(pAys$(), Optional pQ$ = "", Optional pSepChr$ = cCommaSpc)
Dim N%: N = jj.Siz_Ay(pAys): If N% = 0 Then Exit Function
Dim A$: A = jj.Q_S(pAys(0), pQ)
Dim J%
For J = 1 To N - 1
    A = A & pSepChr & jj.Q_S(pAys(J), pQ)
Next
ToStr_Ays = A
End Function
Function ToStr_AyBool$(pAyBool() As Boolean, Optional pQ$ = "", Optional pSepChr$ = cCommaSpc)
Dim N%: N = jj.Siz_Ay(pAyBool): If N% = 0 Then Exit Function
Dim J%: For J = 0 To N - 1
    Dim A$: A = jj.Add_Str(A, jj.Q_S(CStr(pAyBool(J)), pQ), pSepChr)
Next
ToStr_AyBool = A
End Function
Function ToStr_AyByt$(pAyByt() As Byte, Optional pQ$ = "", Optional pSepChr$ = cCommaSpc)
Dim N%: N = jj.Siz_Ay(pAyByt): If N% = 0 Then Exit Function
Dim J%: For J = 0 To N - 1
    Dim A$: A = jj.Add_Str(A, jj.Q_S(CStr(pAyByt(J)), pQ), pSepChr)
Next
ToStr_AyByt = A
End Function
Function ToStr_AyV$(pAyV(), Optional pQ$ = "", Optional pSepChr$ = cCommaSpc)
Dim N%: N = jj.Siz_Ay(pAyV): If N% = 0 Then Exit Function
Dim A$, J%
For J = 1 To N - 1
    If (VarType(pAyV(J)) And vbArray) = 0 Then
        A$ = jj.Add_Str(A$, jj.Q_S(pAyV(J), pQ), pSepChr)
    Else
        Dim mX$: mX = "Array(" & jj.Siz_Ay(pAyV(J)) & ")"
        A$ = jj.Add_Str(A$, jj.Q_S(mX, pQ), pSepChr)
    End If
Next
ToStr_AyV = A$
Exit Function
E: ToStr_AyV = "Err: jj.ToStr_AyV(pAyV).  Msg=" & Err.Description
End Function
Function ToStr_Coll$(pColl As VBA.Collection, Optional pSepChr$ = cComma)
If jj.IsNothing(pColl) Then ToStr_Coll = "#Nothing#": Exit Function
Dim mV, mA$
For Each mV In pColl
    mA = jj.Add_Str(mA, CStr(mV), pSepChr)
Next
ToStr_Coll = mA
End Function
Function ToStr_ColRge$(pCno1 As Byte, pCno2 As Byte)
ToStr_ColRge = jj.Cv_Cno2Col(pCno1) & ":" & jj.Cv_Cno2Col(pCno2)
End Function
Function ToStr_Ctl$(pCtl As Access.Control, Optional pWithTag As Boolean = False)
On Error GoTo R
If pWithTag Then
    If IsNothing(pCtl.Tag) Then
        ToStr_Ctl = pCtl.Name
    Else
        ToStr_Ctl = pCtl.Name & "(" & pCtl.Tag & ")"
    End If
Else
    ToStr_Ctl = pCtl.Name
End If
Exit Function
R: ToStr_Ctl = "Err: jj.ToStr_Ctl(pCtl).  Msg=" & Err.Description
End Function
Function ToStr_Ctls$(pCtls As Access.Controls, Optional pWithTag As Boolean = False, Optional pSepChr$ = cComma)
On Error GoTo R
Dim mS$, iCtl As Access.Control
For Each iCtl In pCtls
    mS = Add_Str(mS, ToStr_Ctl(iCtl, pWithTag), pSepChr)
Next
ToStr_Ctls = mS
Exit Function
R: ToStr_Ctls = "Err: jj.ToStr_Ctls(pCtls).  Msg=" & Err.Description
End Function
Function ToStr_Rel$(pNmRel$, Optional pDb As DAO.Database = Nothing)
On Error GoTo R
Dim mDb As DAO.Database: Set mDb = Cv_Db(pDb)
Dim mRel As DAO.Relation: Set mRel = mDb.Relations(pNmRel)
ToStr_Rel = "Rel(" & pNmRel & "):" & mRel.Table & ";" & mRel.ForeignTable & ";" & jj.ToStr_Flds_Rel(mRel.Fields)
Exit Function
R: ToStr_Rel = "Err: jj.ToStr_Rel(" & pNmRel & ").  Msg=" & Err.Description
End Function
#If Tst Then
Function ToStr_Rel_Tst() As Boolean
Dim mDb As DAO.Database: If jj.Opn_Db_RW(mDb, "C:\Tmp\ProjMeta\Meta\MetaAll.Mdb") Then Stop
Debug.Print jj.ToStr_Rel("AcptR10", mDb)
Shw_DbgWin
End Function
#End If
Function ToStr_Db$(pDb As DAO.Database)
If jj.IsNothing(pDb) Then ToStr_Db = "Nothing": Exit Function
On Error GoTo R
ToStr_Db = pDb.Name
Exit Function
R: ToStr_Db = "Err: jj.ToStr_Db(pDb).  Msg=" & Err.Description
End Function
Function ToStr_FldVal$(pFld As DAO.Field)
On Error GoTo R
ToStr_FldVal = pFld.Value
Exit Function
R: ToStr_FldVal = "#" & Err.Description & "#"
End Function
Function ToStr_Tbl$(pTbl As DAO.TableDef)
Const cSub$ = "ToStr_Tbl"
On Error GoTo R
With pTbl
    ToStr_Tbl = .Name
End With
Exit Function
R: ToStr_Tbl = "Err: jj.ToStr_Tbl(pTbl).  Msg=" & Err.Description
End Function
Function ToStr_Fld$(pFld As DAO.Field, Optional pInclTyp As Boolean = False, Optional pInclVal As Boolean = False)
Const cSub$ = "ToStr_Fld"
On Error GoTo R
With pFld
    If pInclTyp Then
        If pInclVal Then ToStr_Fld = .Name & ":" & jj.ToStr_TypFld(pFld) & "=" & jj.ToStr_FldVal(pFld): Exit Function
        ToStr_Fld = .Name & ":" & jj.ToStr_TypFld(pFld)
        Exit Function
    End If
    If pInclVal Then ToStr_Fld = .Name & "=" & Nz(.ValidateOnSet, "Null"): Exit Function
    ToStr_Fld = .Name
End With
Exit Function
R: ToStr_Fld = "Err: jj.ToStr_Fld(pFld).  Msg=" & Err.Description
End Function
Function ToStr_Fld_Rel$(pFld As DAO.Field)
Const cSub$ = "ToStr_Fld_Rel"
On Error GoTo R
With pFld
    If .Name = .ForeignName Then
        ToStr_Fld_Rel = .Name
    Else
        ToStr_Fld_Rel = .Name & "=" & .ForeignName
    End If
End With
Exit Function
R: ToStr_Fld_Rel = "Err: jj.ToStr_Fld(pFld).  Msg=" & Err.Description
End Function
Function ToStr_Flds_Rel$(pFlds As DAO.Fields, Optional pSepChr$ = cComma)
Const cSub$ = "ToStr_Flds_Rel"
On Error GoTo R
If pFlds.Count = 0 Then ToStr_Flds_Rel = "": Exit Function
Dim mA$, iFld As DAO.Field, J%
For J = 0 To pFlds.Count - 1
    mA = jj.Add_Str(mA, jj.ToStr_Fld_Rel(pFlds(J)), pSepChr)
Next
ToStr_Flds_Rel = mA
Exit Function
R: ToStr_Flds_Rel = "Err: jj.ToStr_Flds(pFlds,pSepChr).  Msg=" & Err.Description
End Function
Function ToStr_Flds$(pFlds As DAO.Fields, Optional pInclTyp As Boolean = False, Optional pInclVal As Boolean = False, Optional pSepChr$ = cComma, Optional pBeg As Byte = 0, Optional pEnd As Byte = 255)
Const cSub$ = "ToStr_Flds"
On Error GoTo R
If pFlds.Count = 0 Then ToStr_Flds = "": Exit Function
Dim mA$, iFld As DAO.Field, J%
For J = pBeg To Fct.MinByt(pFlds.Count - 1, CInt(pEnd))
    mA = jj.Add_Str(mA, jj.ToStr_Fld(pFlds(J), pInclTyp, pInclVal), pSepChr)
Next
ToStr_Flds = mA
Exit Function
R: ToStr_Flds = "Err: jj.ToStr_Flds(pFlds,pInclTypFld,pSepChr,pBeg,pEnd).  Msg=" & Err.Description
End Function
Function ToStr_Flds_Tst() As Boolean
Const cNmt$ = "mstBrand"
Dim mFlds As DAO.Fields: Set mFlds = CurrentDb.TableDefs(cNmt).Fields
Debug.Print jj.ToStr_Flds(CurrentDb.TableDefs(cNmt).Fields, True, True)
jj.Shw_DbgWin
End Function
Function ToStr_Fld_Dcl$(pFld As DAO.Field)
On Error GoTo R
ToStr_Fld_Dcl = pFld.Name & " " & jj.Cv_Fld2Dcl(pFld)
Exit Function
R: ToStr_Fld_Dcl = "Err: jj.ToStr_Fld_Dcl(pFld).  Msg=" & Err.Description
End Function
Function ToStr_Flds_Dcl$(pFlds As DAO.Fields, Optional pSepChr$ = cComma)
Dim mA$
Dim iFld As DAO.Field: For Each iFld In pFlds
    mA = jj.Add_Str(mA, jj.ToStr_Fld_Dcl(iFld), pSepChr)
Next
ToStr_Flds_Dcl = mA
End Function
'Function ToStr_FmRecs(oS$, pSql$, Optional pSep$ = cCommaSpc) As Boolean
'Const cSub$ = "ToStr_FmRecs"
'On Error GoTo R
'oS = ""
'With CurrentDb.OpenRecordset(pSql)
'    While Not .EOF
'        oS = jj.Add_Str(oS, pSep)
'        .MoveNext
'    Wend
'    .Close
'End With
'Exit Function
'R: ss.R
'E: ToStr_FmRecs = True: ss.B cSub, cMod, "pSql", pSql
'End Function
Function ToStr_Frm$(pFrm As Access.Form)
On Error GoTo R
ToStr_Frm = pFrm.Name
Exit Function
R: ToStr_Frm = "Err jj.ToStr_Frm(pFrm).  Msg=" & Err.Description
End Function
Function ToStr_FYNo$(pFyNo As Byte)
ToStr_FYNo = "FY" & Format(pFyNo, "00")
End Function
Function ToStr_Lang$(pLang As eLang)
Select Case pLang
Case eLang.eSC: ToStr_Lang = "SimpChinese"
Case eLang.eTC: ToStr_Lang = "TradChinese"
Case Else: ToStr_Lang = "English"
End Select
End Function
Function ToStr_LpAp$(pSepChr$, pLp$, ParamArray pAp())
Const cSub$ = "ToStr_LpAp"
Dim mAm() As tMap: If jj.Brk_LpVv2Am(mAm, pLp, CVar(pAp)) Then ss.A 1: GoTo E
ToStr_LpAp = jj.ToStr_Am(mAm, , , "[]", pSepChr)
Exit Function
E: ss.C cSub, cMod, "pSepChr,pLp,pAp", pSepChr, pLp, jj.ToStr_Vayv(CVar(pAp))
End Function
#If Tst Then
Function ToStr_LpAp_Tst() As Boolean
Debug.Print jj.ToStr_LpAp(vbLf, "aa,bb,,C", 1, 2, , 1)
End Function
#End If
Function ToStr_Map$(pMap As tMap _
    , Optional pBrkChr$ = "=" _
    , Optional pQ1$ = "" _
    , Optional pQ2$ = "" _
        )
With pMap
    ToStr_Map = jj.Q_S(.F1, pQ1) & pBrkChr & jj.Q_S(.F2, pQ2)
End With
End Function
Function ToStr_Nmt$(pNmt$, Optional pInclTypFld As Boolean = False, Optional pSepChr$ = cComma, Optional pBeg As Byte = 0, Optional pEnd As Byte = 255, Optional pDb As DAO.Database = Nothing)
Const cSub$ = "ToStr_Nmt"
On Error GoTo R
ToStr_Nmt = jj.ToStr_Flds(jj.Cv_Db(pDb).TableDefs(pNmt).Fields, pInclTypFld, , pSepChr, pBeg, pEnd)
Exit Function
R: ToStr_Nmt = "Err: jj.ToStr_Nmt(" & pNmt & cComma & jj.ToStr_Db(pDb) & ").  Msg=" & Err.Description
End Function
Function ToStr_Nmt_Dcl$(pNmt$, Optional pSepChr$ = cComma)
Const cSub$ = "ToStr_Nmt_Dcl"
On Error GoTo R
ToStr_Nmt_Dcl = jj.ToStr_Flds_Dcl(CurrentDb.TableDefs(pNmt).Fields, pSepChr)
Exit Function
R: ss.R
    ToStr_Nmt_Dcl = "Err: jj.ToStr_Nmt_Dcl(" & pNmt & ").  Msg=" & Err.Description
End Function
Function ToStr_Pc$(pPc As PivotCache)
If jj.IsNothing(pPc) Then ToStr_Pc = "#Nothing#": Exit Function
Dim mCmdTxt$: mCmdTxt = "CmdTxt<Nil>"
Dim mCnnStr$: mCnnStr = "CnnStr<Nil>"
Dim mRfhNam$: mRfhNam = "RfhNam<Nil>"
Dim mPcIdx%
On Error Resume Next
With pPc
    mCmdTxt = .CommandText
    mCnnStr = .Connection
    mRfhNam = .RefreshName
    mPcIdx = .Index
End With
On Error GoTo 0
ToStr_Pc = jj.ToStr_LpAp(cComma, "CmdTxt,PcIdx,RfhNam,CnnStr", mCmdTxt, mPcIdx, mRfhNam, mCnnStr)
End Function
Function ToStr_Prp$(pPrp As DAO.Property)
On Error GoTo R
Dim mNm$: mNm = pPrp.Name
ToStr_Prp = mNm & "=[" & pPrp.Value & "]"
Exit Function
R: ss.R
    ToStr_Prp = "Err: jj.ToStr_Prp(" & mNm & ").  Msg=" & Err.Description
End Function
Function ToStr_Prps$(pPrps As DAO.Properties, Optional pSepChr$ = " ")
Dim mA$, J As Byte
On Error GoTo R
For J = 0 To pPrps.Count - 1
    mA = jj.Add_Str(mA, jj.ToStr_Prp(pPrps(J)), pSepChr)
Next
ToStr_Prps = mA
Exit Function
R: ss.R
    ToStr_Prps = "Err: jj.ToStr_Prps(pPrps,pSepChr).  Msg=" & Err.Description
End Function
Function ToStr_Prps_Tst() As Boolean
Dim mPrps As DAO.Properties: Set mPrps = CurrentDb.QueryDefs("ODBCSQry").Properties
Debug.Print jj.ToStr_Prps(mPrps, vbLf)
End Function

Function ToStr_Pt$(pPt As PivotTable)
If jj.IsNothing(pPt) Then ToStr_Pt = "#Nothing#": Exit Function
Dim mCmdTxt$: mCmdTxt = "CmdTxt<Nil>"
Dim mCnnStr$: mCnnStr = "CnnStr<Nil>"
Dim mPcRfhNm$: mPcRfhNm = "PcRfhNm<Nil>"
Dim mPcIdx%
On Error Resume Next
With pPt
    mCmdTxt = .PivotCache.CommandText
    mCnnStr = .PivotCache.Connection
    mPcRfhNm = .PivotCache.RefreshName
    mPcIdx = .PivotCache.Index
End With
On Error GoTo 0
ToStr_Pt = jj.ToStr_LpAp(cComma, "CmdTxt,PcIdx,PtNam,PcRfhNm,CnnStr", mCmdTxt, mPcIdx, pPt.Name, mPcRfhNm, mCnnStr)
End Function
Function ToStr_Qt$(pQt As QueryTable)
If jj.IsNothing(pQt) Then ToStr_Qt = "#Nothing#": Exit Function
Dim mCmdTxt$: mCmdTxt = "CmdTxt<Nil>"
Dim mCnnStr$: mCnnStr = "CnnStr<Nil>"
On Error Resume Next
With pQt
    mCmdTxt = .CommandText
    mCnnStr = .Connection
End With
On Error GoTo 0
ToStr_Qt = jj.ToStr_LpAp(cComma, "CmdTxt,QtNam,CnnStr", mCmdTxt, pQt.Name, mCnnStr)
End Function
Function ToStr_Rge$(pRge As Range)
On Error GoTo R
ToStr_Rge = pRge.Parent.Name & "!" & pRge.Address
Exit Function
R: ToStr_Rge = "Err: jj.ToStr_Rge(pRge).  Msg=" & Err.Description
End Function
Function ToStr_RgeCno$(pRgeCno As tRgeCno)
With pRgeCno
    ToStr_RgeCno = "C" & .Fm & "-" & .To
End With
End Function
Function ToStr_Rs$(pRs As DAO.Recordset, Optional pRsNam$ = "", Optional pSepChr$ = cComma)
Dim mRet$
mRet = "Rs value:" & IIf(pRsNam = "", "", "(RsNam=[" & pRsNam & "])")
On Error GoTo R
Dim iFld As DAO.Field: For Each iFld In pRs.Fields
    With iFld
        mRet = mRet & cComma & jj.ToStr_Fld(iFld)
    End With
Next
ToStr_Rs = mRet
Exit Function
R: ToStr_Rs = "Err: jj.ToStr_Rs(pRs,pRsNam).  Msg=" & Err.Description
End Function
Function ToStr_Rs_NmFld$(pRs As DAO.Recordset, Optional pInclFldCnt As Boolean = False)
On Error GoTo R
Dim mRet$: If pInclFldCnt Then mRet = "NFld(" & pRs.Fields.Count & ") "
Dim iFld As DAO.Field
For Each iFld In pRs.Fields
    mRet = jj.Add_Str(mRet, iFld.Name)
Next
ToStr_Rs_NmFld = mRet
Exit Function
R: ss.R
    ToStr_Rs_NmFld = "Err: jj.ToStr_Rs_NmFld(pRs).  Msg=" & Err.Description
End Function
Function ToStr_Sq$(pSq As tSq)
With pSq
    ToStr_Sq = "(R" & .r1 & ",C" & .c1 & ") - (R" & .r2 & ",C" & .c2 & ")"
End With
End Function
Function ToStr_TypDta$(pTypDta As DAO.DataTypeEnum)
Select Case pTypDta
Case dbBigInt:  ToStr_TypDta = "BigInt": Exit Function
Case dbBinary:  ToStr_TypDta = "Binary": Exit Function
Case dbBoolean: ToStr_TypDta = "YesNo":   Exit Function
Case dbByte:    ToStr_TypDta = "Byte":   Exit Function
Case dbChar:    ToStr_TypDta = "Char": Exit Function
Case dbCurrency: ToStr_TypDta = "Currency": Exit Function
Case dbDate:    ToStr_TypDta = "Date": Exit Function
Case dbDecimal: ToStr_TypDta = "Decimal": Exit Function
Case dbDouble:  ToStr_TypDta = "Double": Exit Function
Case dbFloat:   ToStr_TypDta = "Float": Exit Function
Case dbGUID:    ToStr_TypDta = "GUID": Exit Function
Case dbInteger: ToStr_TypDta = "Int": Exit Function
Case dbLong:    ToStr_TypDta = "Long": Exit Function
Case dbLongBinary: ToStr_TypDta = "LongBinary": Exit Function
Case dbMemo:    ToStr_TypDta = "Memo": Exit Function
Case dbNumeric: ToStr_TypDta = "Numeric": Exit Function
Case dbSingle:  ToStr_TypDta = "Single": Exit Function
Case dbText:    ToStr_TypDta = "Text": Exit Function
Case dbTime:    ToStr_TypDta = "Time": Exit Function
Case dbTimeStamp: ToStr_TypDta = "TimeStamp":  Exit Function
Case dbVarBinary: ToStr_TypDta = "VarBinary":    Exit Function
Case Else:      ToStr_TypDta = "Unknow FieldTyp(" & pTypDta & ")"
End Select
End Function
Function ToStr_TypFld$(pFld As DAO.Field)
With pFld
    If .Type = dbText Then
        ToStr_TypFld = jj.ToStr_TypDta(.Type) & .Size
    Else
        ToStr_TypFld = jj.ToStr_TypDta(.Type)
    End If
End With
End Function
Function ToStr_TypMsg$(pTypMsg As eTypMsg)
Select Case pTypMsg
Case eTypMsg.ePrmErr:    ToStr_TypMsg = "PrmErr"
Case eTypMsg.eCritical:  ToStr_TypMsg = "Critical"
Case eTypMsg.eTrc:       ToStr_TypMsg = "Trace"
Case eTypMsg.eWarning:   ToStr_TypMsg = "Warning"
Case eTypMsg.eSeePrvMsg: ToStr_TypMsg = "SeePrvMsg"
Case eTypMsg.eException: ToStr_TypMsg = "Exception"
Case eTypMsg.eUsrInfo:   ToStr_TypMsg = "User Information"
Case eTypMsg.eRunTimErr: ToStr_TypMsg = "RunTimErr"
Case eTypMsg.eImpossibleReachHere: ToStr_TypMsg = "ImpossibleReachHere"
Case eTypMsg.eQuit: ToStr_TypMsg = "Application Quit"
Case Else: ToStr_TypMsg = "??(" & pTypMsg & ")"
End Select
'    ePrmErr = 1
'    eCritical = 2
'    eTrc = 3
'    eWarning = 4
'    eSeePrvMsg = 5
'    eException = 6
'    eUsrInfo = 7
'    eRunTimErr = 8
'    eImpossibleReachHere = 9
End Function
Function ToStr_TypObj$(pTypObj As AcObjectType)
Select Case pTypObj
Case AcObjectType.acForm:    ToStr_TypObj = "Forms":     Exit Function
Case AcObjectType.acQuery:   ToStr_TypObj = "Queries":   Exit Function
Case AcObjectType.acTable:   ToStr_TypObj = "Tables":    Exit Function
Case AcObjectType.acReport:  ToStr_TypObj = "Reports":   Exit Function
End Select
ToStr_TypObj = "AcObjectType(" & pTypObj & ")"
End Function
Function ToStr_TypQry$(pTypQry As DAO.QueryDefTypeEnum)
Select Case pTypQry
Case DAO.QueryDefTypeEnum.dbQAction:    ToStr_TypQry = "Action"
Case DAO.QueryDefTypeEnum.dbQAppend:    ToStr_TypQry = "Append"
Case DAO.QueryDefTypeEnum.dbQCompound:  ToStr_TypQry = "Compound"
Case DAO.QueryDefTypeEnum.dbQCrosstab:  ToStr_TypQry = "Crosstab"
Case DAO.QueryDefTypeEnum.dbQDDL:       ToStr_TypQry = "DDL"
Case DAO.QueryDefTypeEnum.dbQDelete:    ToStr_TypQry = "DDL"
Case DAO.QueryDefTypeEnum.dbQMakeTable: ToStr_TypQry = "MakeTable"
Case DAO.QueryDefTypeEnum.dbQProcedure: ToStr_TypQry = "Procedure"
Case DAO.QueryDefTypeEnum.dbQSelect:    ToStr_TypQry = "Select"
Case DAO.QueryDefTypeEnum.dbQSetOperation:  ToStr_TypQry = "SetOperation"   'Union
Case DAO.QueryDefTypeEnum.dbQSPTBulk:       ToStr_TypQry = "SPTBulk"
Case DAO.QueryDefTypeEnum.dbQSQLPassThrough: ToStr_TypQry = "SqlPassThrough"
Case DAO.QueryDefTypeEnum.dbQUpdate:        ToStr_TypQry = "Update"
Case Else: ToStr_TypQry = "Unknown(" & pTypQry & ")"
End Select
End Function
Function ToStr_TblAtr$(pTblAtr&)
Dim mA$
If pTblAtr And DAO.TableDefAttributeEnum.dbAttachedODBC Then mA = "ODBC"
If pTblAtr And DAO.TableDefAttributeEnum.dbAttachedTable Then mA = jj.Add_Str(mA, "Lnk", " ")
If pTblAtr And DAO.TableDefAttributeEnum.dbHiddenObject Then mA = jj.Add_Str(mA, "Hide", " ")
If pTblAtr And DAO.TableDefAttributeEnum.dbSystemObject Then mA = jj.Add_Str(mA, "Sys", " ")
ToStr_TblAtr = mA
End Function
Function ToStr_V$(pV As Variant)    ' pV is array of variant
Const cSub$ = "ToStr_V"
If VarType(pV) <> vbArray + vbVariant Then ss.A 1, "pV must be VarTyp of Array+Var", , "VarTyp of pV", VarType(pV): GoTo E
Dim mAyV(): mAyV = pV
Dim A$: A$ = mAyV(0)
Dim J As Byte
For J = 1 To UBound(mAyV)
    A$ = A$ & cComma & mAyV(J)
Next
ToStr_V = A$
Exit Function
E:
: ss.B cSub, cMod, "pV", pV
    ToStr_V = "Err: jj.ToStr_V(pV).  Msg=" & Err.Description
End Function
Function ToStr_VBPrj$(pVBPrj As VBProject)
On Error GoTo R
ToStr_VBPrj = pVBPrj.Name
Exit Function
R: ss.R
    ToStr_VBPrj = "Err: jj.ToStr_VBPrj(pVbPrj).  Msg=" & Err.Description
End Function
Function ToStr_Wb$(pWb As Workbook)
On Error GoTo R
ToStr_Wb = pWb.FullName
Exit Function
R: ss.R
    ToStr_Wb = "Err: jj.ToStr_Wb(pWb).  Msg=" & Err.Description
End Function
Function ToStr_Wrd$(pWrd As Word.Document)
On Error GoTo R
ToStr_Wrd = pWrd.FullName
Exit Function
R: ss.R
    ToStr_Wrd = "Err: jj.ToStr_Wrd(pWrd).  Msg=" & Err.Description
End Function
Function ToStr_Ws$(pWs As Worksheet, Optional pInclNmWb As Boolean = False)
On Error GoTo R
If pInclNmWb Then ToStr_Ws = "Wb=" & jj.ToStr_Wb(pWs.Parent) & ", Ws=" & pWs.Name: Exit Function
ToStr_Ws = "Ws=" & pWs.Name
Exit Function
R: ss.R
    ToStr_Ws = "Err: jj.ToStr_Ws(pWs,pInclWb).  Msg=" & Err.Description
End Function
Function ToStr_YrWk$(pYr As Byte, pWk As Byte)
ToStr_YrWk = "Yr" & Format(pYr, "00") & "_Wk" & Format(pWk, "00")
End Function

