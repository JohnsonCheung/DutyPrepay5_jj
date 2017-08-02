Attribute VB_Name = "xImp"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xImp"
Function Imp_FndCurVal(oCurVal, pTblImp&, pFldImp&, pNmTblImp$, pNmFldImp$, pId&) As Boolean
'Aim: From [$ImpTblF] find the way the of how to find the current value by [$ImpTblF]->TypImpCurVal.  If no value find, use query.
Const cSub$ = "Imp_FndCurVal"
On Error GoTo R
Dim mTypImpCurVal As jj.eTypImpCurVal, mTblRf&, mFldRf&
Dim mSql$
mSql = "Select TypImpCurVal,TblRf,FldRf from ([$ImpTblF] where tf.Tbl=" & pTblImp & " and tf.Fld=" & pFldImp
If jj.Fnd_ValFmSql3(mTypImpCurVal, mTblRf, mFldRf, mSql) Then
    If jj.Run_Qry("qryImpFndCurVal#" & pNmTblImp & "#" & pNmFldImp, , , True, "Id=" & pId) Then ss.A 1: GoTo E
    If jj.Fnd_ValFmSql(oCurVal, "Select CurVal from [@CurVal]") Then ss.A 2, "Err in Find [@CurVal] after running qry[qryImpFndCurVal#<pNmTblImp>#<pNmFldImp>]": GoTo E
    Exit Function
End If
Select Case mTypImpCurVal
Case eTypImpCurVal.eTblF
    Dim mNmTblRf$, mNmFldRf$
    If jj.Fnd_Nm_ById(mNmTblRf, "Tbl", mTblRf) Then ss.A 3: GoTo E
    If jj.Fnd_Nm_ById(mNmFldRf, "Fld", mFldRf) Then ss.A 3: GoTo E
    mSql = jj.Fmt_Str("Select {0} from [${1}] where {0}={2}", mNmTblRf, mNmFldRf, pId)
    If jj.Fnd_ValFmSql(oCurVal, mSql) Then ss.A 2: GoTo E
    Exit Function
End Select
ss.A 1, "Unexpected TypImpCurVal", , "TypImpCurVal of pTblImp,pFldImp", mTypImpCurVal: GoTo E
Exit Function
R: ss.R
E: Imp_FndCurVal = True: ss.B cSub, cMod, "pTblImp,pFldImp,pNmTblImp,pNmFldImp,pId", pTblImp, pFldImp, pNmTblImp, pNmFldImp, pId
End Function
Function Imp_VdtFld(oIsChg As Boolean, oVdtMsg$, pTblImp&, pNmTblImp$, pNmFldImp$, pId&, pV) As Boolean
'Aim: All fields in the table {pNmTblImp} should be all be valid before the table is updated.  This routine is vdt one field at a time.
'     How to validate is user-definible from MetaDb
Const cSub$ = "ImpVdt"
On Error GoTo R
Dim mFldImp&: If jj.Fnd_Id_ByNm(mFldImp, "Fld", pNmFldImp) Then ss.A 1: GoTo E
Dim mCurVal: If jj.Imp_FndCurVal(mCurVal, pTblImp, mFldImp, pNmTblImp, pNmFldImp, pId) Then ss.A 1: GoTo E
oIsChg = mCurVal <> pV
If Not oIsChg Then Exit Function

Dim mVdt&
If jj.Fnd_ValFmSql(mVdt, "Select Vdt from [$TblF]" & _
    " where Tbl in (Select Tbl from [$Tbl] where NmTbl='" & pNmTblImp & "')" & _
    " and   Fld in (Select Fld from [$Fld] where NmFld='" & pNmFldImp & "'") Then ss.A 1, "Given pNmTblImp & pNmFldImp not find in $Tbl,$Fld & $TblF": GoTo E
Exit Function
If jj.Vdt_Fld(oVdtMsg, pId&, pV, mVdt) Then ss.A 2: GoTo E
Exit Function
R: ss.R
E: Imp_VdtFld = True: ss.B cSub, cMod, "pNmTblImp,pNmFldImp,pId,pV", pNmTblImp, pNmFldImp, pId, pV
End Function
Function Imp_Ty(pItm$, Optional pMaxTy As Byte = 1, Optional pNmtImp$ = "") As Boolean
'Aim: It is required to import all records in the import table [>{pItm}] into the Ty tables.
'     Assume there are tables in currentDb:
'       Itm Table:   name = [$<pItm>]       Example, $Tbl  (MaxTy=2)
'       Import Table:name = [><pItm>]       Example, [>Tbl]  (Note: if pNmtImp is given, use pNmtImp, eg. >#Tbl
'       Ty Tables:   name = [$Ty<pItm>]     Example, $TyTbl for each record in $Tbl.  3 fields: Tbl, TyTbl1, TyTbl2
'       Ty1 Tables:  name = [$Ty<pItm>1]    Example, $TyTbl1.                         4 fields: TyTbl1,    NmTyTbl1,    TyTbl1x,   DesTyTbl1
'                           [$Ty<pItm>1x]   Example, $TyTbl1x                         4 fields: TyTbl1x,   NmTyTbl1x,   TyTbl1xx,  DesTyTbl1
'                           [$Ty<pItm>1xx]  Example, $TyTbl1xx                        4 fields: TyTbl1xx,  NmTyTbl1xx,  TyTbl1xxx, DesTyTbl1
'                           [$Ty<pItm>1xxx] Example, $TyTbl1xxx                       3 fields: TyTbl1xxx, NmTyTbl1xxx,            DesTyTbl1
'       Ty2 Tables:  name = [$Ty<pItm>2]    Example, $TyTbl2.                         4 fields: TyTbl2,    NmTyTbl2,    TyTbl2x,   DesTyTbl2
'                           [$Ty<pItm>2x]   Example, $TyTbl2x                         4 fields: TyTbl2x,   NmTyTbl2x,   TyTbl2xx,  DesTyTbl2
'                           [$Ty<pItm>2xx]  Example, $TyTbl2xx                        4 fields: TyTbl2xx,  NmTyTbl2xx,  TyTbl2xxx, DesTyTbl2
'                           [$Ty<pItm>2xxx] Example, $TyTbl2xxx                       3 fields: TyTbl2xxx, NmTyTbl2xxx,            DesTyTbl2
'     Assume in the table [$<pItm>] has following fields:
'       <pItm> & Nm<pItm>               Example, Table [$Tbl] will have 2 fields: [Tbl] & [NmTbl]
'     Assume in there is import table named as [><pNmtImp>], example, [>Tbl] has following fields:
'       Nm<pItm>, and,
'       NmTy<pItm>1, NmTy<pItm>1x, NmTy<pItm>1xx, NmTy<pItm>1xxx, and,
'       NmTy<pItm>2, NmTy<pItm>2x, NmTy<pItm>2xx, NmTy<pItm>2xxx.
'Logic:
'Check all tables & fields are correct.
'Check each Ty must in tree (no child belongs to 2 parents)
'Add dummy rec to $<Itm>Ty{N}{x} ({N}=1-pMaxTy, {x}=x,..,xxx): $<Itm>Ty{N}{x}: <Itm>Ty{N}{x}, <Itm>Ty{N}{x}x, Des<Itm>Ty{N}{x}, Nm<Itm>Ty{N}{x}
'Build #Nm<pItm>Tbl: NmTyTbl1, .., NmTyTbl1xxx from >Tbl
'For B=1 to pMaxTy
'    Build Table Ty<J>xxx
'    Build Table Ty<J>xx
'    Build Table Ty<J>x
'Next
'Build Tbl Ty
Const cSub$ = "Imp_Ty"
If jj.Dlt_Tbl_ByPfx("#") Then ss.A 1: GoTo E

'Check all tables & fields are correct.
Dim mNmtImp1$: If pNmtImp = "" Then mNmtImp1 = jj.Q_S(pItm, "[>*]") Else mNmtImp1 = jj.Q_SqBkt(pNmtImp)
Dim mNmtItm$: mNmtItm = jj.Q_S(pItm, "[$*]")
Dim mNmtTy$: mNmtTy = jj.Q_S(pItm, "[$Ty*]")
ReDim mNmtTyN$(pMaxTy - 1), mNmtTyNx$(pMaxTy - 1), mNmtTyNxx$(pMaxTy - 1), mNmtTyNxxx$(pMaxTy - 1)
Dim B As Byte
For B = 0 To pMaxTy - 1
    mNmtTyN(B) = jj.Q_S(pItm, jj.Fmt_Str("[$Ty*{0}]", B + 1))
    mNmtTyNx(B) = jj.Q_S(pItm, jj.Fmt_Str("[$Ty*{0}x]", B + 1))
    mNmtTyNxx(B) = jj.Q_S(pItm, jj.Fmt_Str("[$Ty*{0}xx]", B + 1))
    mNmtTyNxxx(B) = jj.Q_S(pItm, jj.Fmt_Str("[$Ty*{0}xxx]", B + 1))
Next

Dim mLm$
mLm = jj.Fmt_Str("Itm={0};N={1};X=,x,xx,xxx", pItm, jj.Fmt_Str_Repeat(1, pMaxTy))
'Chk Tbl Exist
Dim mAnt$(): If jj.Fmt_Str_Repeat_ByLm_IntoAy(mAnt, "$Ty{Itm}{N}{X}", mLm) Then ss.A 1: GoTo E
If Not jj.IsLnt(Join(mAnt, ",")) Then ss.A 2: GoTo E
If Not jj.IsTbl("$Ty" & pItm) Then ss.A 3: GoTo E
If Not jj.IsTbl(mNmtItm) Then ss.A 4: GoTo E
If Not jj.IsTbl(mNmtImp1) Then ss.A 5: GoTo E
Do
    Dim mA$, mB$
    mA = mNmtItm: If jj.Chk_Struct_Tbl_SubSet(mA, jj.Fmt_Str("{0}, Nm{0}", pItm)) Then ss.A 3: GoTo E
    mA = mNmtImp1: If jj.Chk_Struct_Tbl_SubSet(mA, jj.Bld_Struct_ForTy_Import(pItm, pMaxTy)) Then ss.A 6: GoTo E
    mA = mNmtTy:  If jj.Chk_Struct_Tbl_SubSet(mA, jj.Bld_Struct_ForTy(pItm, , CStr(pMaxTy))) Then ss.A 9: GoTo E
    For B = 0 To pMaxTy - 1
        mA = mNmtTyN(B):    If jj.Chk_Struct_Tbl_SubSet(mA, jj.Bld_Struct_ForTy(pItm, B + 1)) Then ss.A 12: GoTo E
        mA = mNmtTyNx(B):   If jj.Chk_Struct_Tbl_SubSet(mA, jj.Bld_Struct_ForTy(pItm, B + 1, "x")) Then ss.A 15: GoTo E
        mA = mNmtTyNxx(B):  If jj.Chk_Struct_Tbl_SubSet(mA, jj.Bld_Struct_ForTy(pItm, B + 1, "xx")) Then ss.A 18: GoTo E
        mA = mNmtTyNxxx(B): If jj.Chk_Struct_Tbl_SubSet(mA, jj.Bld_Struct_ForTy(pItm, B + 1, "xxx")) Then ss.A 21: GoTo E
    Next
Loop Until True
'Check each Ty must in tree (no child belongs to 2 parents)
Do
    '     Assume in there is import table named as {mNmtImp1}, example, {mNmtImp1} has following fields:
    '       Nm<pItm>, and,
    '       NmTy<pItm>1, NmTy<pItm>1x, NmTy<pItm>1xx, NmTy<pItm>1xxx, and,
    '       NmTy<pItm>2, NmTy<pItm>2x, NmTy<pItm>2xx, NmTy<pItm>2xxx.
    Dim mNmFldChd$, mNmFldPar$, mPfx$: mPfx = "NmTy" & pItm
    For B = 1 To pMaxTy
        mNmFldChd = mPfx & B
        mNmFldPar = mPfx & B & "x"
        If jj.Chk_No2Par(mNmtImp1, mNmFldChd, mNmFldPar) Then ss.A 22: GoTo E
        mNmFldChd = mPfx & B & "x"
        mNmFldPar = mPfx & B & "xx"
        If jj.Chk_No2Par(mNmtImp1, mNmFldChd, mNmFldPar) Then ss.A 23: GoTo E
        mNmFldChd = mPfx & B & "xx"
        mNmFldPar = mPfx & B & "xxx"
        If jj.Chk_No2Par(mNmtImp1, mNmFldChd, mNmFldPar) Then ss.A 24: GoTo E
    Next
Loop Until True
'Add dummy rec to $<Itm>Ty{N}{x} ({N}=1-pMaxTy, {x}=x,..,xxx)
'       $<Itm>Ty{N}{x}: <Itm>Ty{N}{x}, <Itm>Ty{N}{x}x, Des<Itm>Ty{N}{x}, Nm<Itm>Ty{N}{x}
Dim mNmt$:   mNmt = jj.Q_S(pItm, "$Ty*1")
Dim mCndn$: mCndn = jj.Q_S(pItm & "Ty*1=0")
Dim mRecCnt&: If jj.Fnd_RecCnt_ByNmtq(mRecCnt, mNmt, mCndn) Then ss.A 25: GoTo E
Dim mSql$
If mRecCnt = 0 Then
    For B = 1 To pMaxTy
        mSql = jj.Fmt_Str("Insert into [$Ty{0}{1}xxx] (Ty{0}{1}xxx, NmTy{0}{1}xxx ) values (0,'-')", pItm, B)
                If jj.Run_Sql(mSql) Then ss.A 26: GoTo E
        mSql = jj.Fmt_Str("Insert into [$Ty{0}{1}xx] (Ty{0}{1}xx, Ty{0}{1}xxx, NmTy{0}{1}xx ) values (0,0,'-')", pItm, B)
                If jj.Run_Sql(mSql) Then ss.A 27: GoTo E
        mSql = jj.Fmt_Str("Insert into [$Ty{0}{1}x] (Ty{0}{1}x, Ty{0}{1}xx, NmTy{0}{1}x ) values (0,0,'-')", pItm, B)
                If jj.Run_Sql(mSql) Then ss.A 28: GoTo E
        mSql = jj.Fmt_Str("Insert into [$Ty{0}{1}] (Ty{0}{1}, Ty{0}{1}x, NmTy{0}{1} ) values (0,0,'-')", pItm, B)
                If jj.Run_Sql(mSql) Then ss.A 29: GoTo E
    Next
End If

'Build table [#NmTy<pItm><B>]: NmTyTbl<B>, .., NmTyTbl<B>xxx from {mNmtImp1}
Dim mAySql$()

Dim mFmtStr$
'mFmtStr = "SELECT DISTINCT [Imp].NmTyTbl2, [Imp].NmTyTbl2x, [Imp].NmTyTbl2xx, [Imp].NmTyTbl2xxx" & _
'" INTO [#NmTyTbl2]" & _
'" FROM [>Tbl] AS Imp;"
mFmtStr = "SELECT DISTINCT [Imp].NmTy{Itm}{N}, [Imp].NmTy{Itm}{N}x, [Imp].NmTy{Itm}{N}xx, [Imp].NmTy{Itm}{N}xxx" & _
" INTO [#NmTy{Itm}{N}]" & _
" FROM {NmtImp1} AS Imp;"
mLm = jj.Fmt_Str("Itm={0};N={1};X=xx,x,;NmtImp1={2}", pItm, jj.Fmt_Str_Repeat(1, pMaxTy), mNmtImp1)
If jj.Run_Sql_By_Repeat_ByLm(mFmtStr, mLm) Then ss.A 30: GoTo E

'mFmtStr = "Update [#NmTyTbl1] set [NmTyTbl1x]=[NmTyTbl1] where [NmTyTbl1x] is Null"
mFmtStr = "Update [#NmTy{Itm}{N}] set [NmTy{Itm}{N}{X}x]=[NmTy{Itm}{N}{X}] where [NmTy{Itm}{N}{X}x] is Null"
If jj.Run_Sql_By_Repeat_ByLm(mFmtStr, mLm) Then ss.A 30: GoTo E

'For B=1 to pMaxTy
'    Build Table Ty<J>xxx
'    Build Table Ty<J>xx
'    Build Table Ty<J>x
'Next
'Build Tbl Ty

'mFmtStr = "INSERT INTO [$TyTbl2xxx] ( NmTyTbl2xxx )" & _
" SELECT Src.NmTyTbl2xxx" & _
" FROM [#NmTyTbl2] AS Src LEFT JOIN [$TyTbl2xxx] AS Tar ON Src.NmTyTbl2xxx = Tar.NmTyTbl2xxx" & _
" Where (((Tar.NmTyTbl2xxx) Is Null) And ((Nz([NmTyTbl2xxx], "")) <> ""))" & _
" GROUP BY Src.NmTyTbl2xxx;
'mFmtStr = "INSERT INTO [$TyTbl{N}xxx] ( NmTyTbl{N}xxx )" & _
'" SELECT Src.NmTyTbl{N}xxx" & _
'" FROM [#NmTyTbl{N}] AS Src LEFT JOIN [$TyTbl{N}xxx] AS Tar ON Src.NmTyTbl{N}xxx = Tar.NmTyTbl{N}xxx" & _
'" Where (((Tar.NmTyTbl{N}xxx) Is Null) And ((Nz(Src.[NmTyTbl{N}xxx], '')) <> ''))" & _
'" GROUP BY Src.NmTyTbl{N}xxx;
mFmtStr = "INSERT INTO [$TyTbl{N}xxx] ( NmTyTbl{N}xxx )" & _
" SELECT Src.NmTyTbl{N}xxx" & _
" FROM [#NmTyTbl{N}] AS Src LEFT JOIN [$TyTbl{N}xxx] AS Tar ON Src.NmTyTbl{N}xxx = Tar.NmTyTbl{N}xxx" & _
" Where (((Tar.NmTyTbl{N}xxx) Is Null) And ((Nz(Src.[NmTyTbl{N}xxx], '')) <> ''))" & _
" GROUP BY Src.NmTyTbl{N}xxx;"
mLm = "Itm=Tbl;N=" & jj.Fmt_Str_Repeat(1, pMaxTy)
If jj.Run_Sql_By_Repeat_ByLm(mFmtStr, mLm) Then ss.A 30: GoTo E


mFmtStr = "INSERT INTO [$TyTbl{N}{x}] ( NmTyTbl{N}{x}, TyTbl{N}{x}x )" & _
" SELECT Src.NmTyTbl{N}{x}, Par.TyTbl{N}{x}x" & _
" FROM ([#NmTyTbl{N}] AS Src LEFT JOIN [$TyTbl{N}{x}] AS Tar ON Src.NmTyTbl{N}{x} = Tar.NmTyTbl{N}{x}) INNER JOIN [$TyTbl{N}{x}x] AS Par ON Src.NmTyTbl{N}{x}x = Par.NmTyTbl{N}{x}x" & _
" GROUP BY Src.NmTyTbl{N}{x}, Par.TyTbl{N}{x}x, Tar.NmTyTbl{N}{x}" & _
" HAVING (((Tar.NmTyTbl{N}{x}) Is Null));"
mLm = jj.Fmt_Str("Itm={0};N={1};X=xx,x,", pItm, jj.Fmt_Str_Repeat(1, pMaxTy))
If jj.Run_Sql_By_Repeat_ByLm(mFmtStr, mLm) Then ss.A 30: GoTo E

'Build $TyTbl from $Tbl, $TyTbl{N}
'       $TyTbl: Tbl, TyTbl{N}
''Make #TyTbl
'''SELECT [$Tbl].Tbl, Nz(Ty1.TyTbl1,0) AS TyTbl1, Nz(Ty2.TyTbl2,0) AS TyTbl2
''' INTO [#TyTbl]
''' FROM (([>#Tbl] AS Imp
''' INNER JOIN [$Tbl] ON [Imp].NmTbl = [$Tbl].NmTbl)
''' LEFT JOIN [$TyTbl1] AS Ty1 ON [Imp].NmTyTbl1 = Ty1.NmTyTbl1)
''' LEFT JOIN [$TyTbl2] AS Ty2 ON [Imp].NmTyTbl2 = Ty2.NmTyTbl2
''' WHERE (Nz(Ty1.TyTbl1,0)<>0) OR (Nz(Ty2.TyTbl2,0)<>0)

'''SELECT [${Itm}].{Itm}, Nz(Ty1.Ty{Itm}1,0) AS Ty{Itm}1, Nz(Ty2.Ty{Itm}2,0) AS Ty{Itm}2
''' INTO [#Ty{Itm}]
''' FROM (({NmtImp1} Imp
''' INNER JOIN [${Itm}] ON [Imp].Nm{Itm} = [${Itm}].Nm{Itm})
''' LEFT JOIN [$Ty{Itm}1] AS Ty1 ON [Imp].NmTy{Itm}1 = Ty1.NmTy{Itm}1)
''' LEFT JOIN [$Ty{Itm}2] AS Ty2 ON [Imp].NmTy{Itm}2 = Ty2.NmTy{Itm}2
''' WHERE (((Nz(Ty1.Ty{Itm}1,0))<>0)) OR (((Nz(Ty2.Ty{Itm}2,0))<>0));
mA = jj.Fmt_Str("Nz(Ty{N}.[Ty{0}{N}],0) As Ty{0}{N}", pItm):
    Dim mLst$:      mLst = jj.Fmt_Str_Repeat(1, pMaxTy, mA)
mA = jj.Fmt_Str(" LEFT JOIN [$Ty{0}{N}] Ty{N} ON [Imp].NmTy{0}{N} = Ty{N}.NmTy{0}{N})", pItm, mNmtImp1)
    Dim mLeftJoin$: mLeftJoin$ = jj.Fmt_Str_Repeat(1, pMaxTy, mA, "")
Dim mBracket$:  mBracket = String(pMaxTy, "(")
mSql = jj.Fmt_Str("SELECT [${0}].{0}, {1}" & _
    " INTO [#Ty{0}]" & _
    " FROM {3}({4} Imp" & _
    " INNER JOIN [${0}] ON [Imp].Nm{0} = [${0}].Nm{0})" & _
    " {2}" _
    , pItm, mLst, mLeftJoin, mBracket, mNmtImp1)
If jj.Run_Sql(mSql) Then ss.A 38: GoTo E

''Append #TyTbl to $TyTbl
'''Insert Into [$TyTbl]
''' Select tmp.Tbl, Tmp.TyTbl1,Tmp.TyTbl2
''' From [#TyTbl] Tmp Left Join [$TyTbl] Ty on Tmp.Tbl = Ty.Tbl
''' Where Ty.Tbl Is Null
mA = jj.Fmt_Str("Tmp.Ty{0}{N}", pItm):
mLst = jj.Fmt_Str_Repeat(1, pMaxTy, mA)
mSql = jj.Fmt_Str("Insert" & _
    " Into [$Ty{0}]" & _
    " Select tmp.{0}, {1}" & _
    " From [#Ty{0}] Tmp Left Join [$Ty{0}] Ty on Tmp.{0} = Ty.{0}" & _
    " Where Ty.{0} Is Null" _
    , pItm, mLst)
If jj.Run_Sql(mSql) Then ss.A 39: GoTo E

''Update from #TyTbl
'''Update [#TyTbl] Tmp
''' Inner Join [$TyTbl] Ty on Tmp.Tbl = Ty.Tbl
''' Set Ty.TyTbl1=Tmp.TyTbl1
'''   , Ty.TyTbl2=Tmp.TyTbl2
mA = jj.Fmt_Str("Ty.Ty{0}{N}=Tmp.Ty{0}{N}", pItm)
mLst = jj.Fmt_Str_Repeat(1, pMaxTy, mA)
mSql = jj.Fmt_Str("Update [#Ty{0}] Tmp" & _
    " Inner Join [$Ty{0}] Ty on Tmp.{0} = Ty.{0}" & _
    " Set {1}" _
    , pItm, mLst)
If jj.Run_Sql(mSql) Then ss.A 40: GoTo E
GoTo X
R: ss.R
E: Imp_Ty = True: ss.B cSub, cMod, "pItm,pMaxTy", pItm, pMaxTy
X:
End Function
#If Tst Then
Function Imp_Ty_Tst() As Boolean
Dim mFbPgm$: mFbPgm = "p:\workingdir\PgmObj\JMtcDb.mdb"
Dim mFbDta$: mFbDta = "p:\workingdir\MetaDb.mdb"
If False Then
    If jj.Crt_Tbl_FmLnkNmt(mFbPgm, ">#Tbl") Then Stop
    If jj.Crt_Tbl_FmLnkNmt(mFbDta, "$Tbl") Then Stop
    If jj.Crt_Tbl_FmLnkSetNmt(mFbDta, "$TyTbl*") Then Stop
End If
If jj.Run_Sql("Delete * from [$TyTbl]") Then Stop
If jj.Run_Sql("Delete * from [$TyTbl1] where TyTbl1<>0") Then Stop
If jj.Run_Sql("Delete * from [$TyTbl1x] where TyTbl1x<>0") Then Stop
If jj.Run_Sql("Delete * from [$TyTbl1xx] where TyTbl1xx<>0") Then Stop
If jj.Run_Sql("Delete * from [$TyTbl1xxx] where TyTbl1xxx<>0") Then Stop
If jj.Run_Sql("Delete * from [$TyTbl2] where TyTbl2<>0") Then Stop
If jj.Run_Sql("Delete * from [$TyTbl2x] where TyTbl2x<>0") Then Stop
If jj.Run_Sql("Delete * from [$TyTbl2xx] where TyTbl2xx<>0") Then Stop
If jj.Run_Sql("Delete * from [$TyTbl2xxx] where TyTbl2xxx<>0") Then Stop
If jj.Imp_Ty("Tbl", 2, ">#Tbl") Then Stop
End Function
#End If
