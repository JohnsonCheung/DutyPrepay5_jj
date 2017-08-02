Attribute VB_Name = "xBrk"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xBrk"
Private Enum eSts
ePub_Prv = 1
ePrp_Fct_Sub = 3
eGet_Set_Let = 4
eNmPrc = 5
eArg = 6
eAim = 7
EEnd = 99
End Enum
Function Brk_FldDcl(oNmFld$, oTyp As DAO.DataTypeEnum, oLen As Byte, pFldDcl$) As Boolean
Const cSub$ = "Brk_FldDlc"
On Error GoTo R
Dim mFldDcl$: mFldDcl = Trim(pFldDcl)
Dim mP%: mP = InStr(mFldDcl, " ")
If mP = 0 Then ss.A 1, "There is not space in pFldDcl": GoTo E
oNmFld = Replace(Left(mFldDcl, mP - 1), "^", " ")
If jj.Cv_TypDAO_FmFldDcl(oTyp, oLen, mID(mFldDcl, mP + 1)) Then ss.A 2: GoTo E
Exit Function
R: ss.R
E: Brk_FldDcl = True: ss.B cSub, cMod, "pFldDcl", pFldDcl
End Function
Function Brk_Rec_ByCmd(pBrkRecCmd$, Optional pSepChr$ = cComma) As Boolean
'Aim: run BrkRec by pBrkRecCmd, which is a format:
'     Brk Split [To] [Into] [Keep] [SetSno] [Stp]
Const cSub$ = "Brk_Rec_ByCmd"
Dim mBrk$, mSplit$, mTo$, mInto$, mKeep$, mSetSno$, mBeg%, mStp%
If jj.Brk_Brk_RecCmd(mBrk, mSplit, mInto, mTo, mKeep, mSetSno, mBeg, mStp, pBrkRecCmd) Then ss.A 1: GoTo E
If jj.Brk_Rec(mBrk, mSplit, mInto, mTo, mKeep, mSetSno, mBeg, mStp, pSepChr) Then ss.A 2: GoTo E
Exit Function
E: Brk_Rec_ByCmd = True: ss.B cSub, cMod, "pBrkRecCmd$", pBrkRecCmd
End Function
#If Tst Then
Function Brk_Rec_ByCmd_Tst() As Boolean
Const cBrk$ = "#Tmp"
'Create cBrk
If jj.Dlt_Tbl(cBrk) Then Stop: GoTo E
Dim mSql$: mSql = "Create table [" & cBrk & "] (abc text(10), def integer, xxx text(255))"
If jj.Run_Sql(mSql) Then Stop
Dim J%
For J = 0 To 9
    Dim mA$: mA = "aa," & J & cComma & J + 1 & cComma & J + 2
    mSql = jj.Fmt_Str("Insert into [{0}] values ('abc-{1}',{2},'{3}')", cBrk, J, J * 2, mA)
    If jj.Run_Sql(mSql) Then Stop
Next
If jj.Brk_Rec_ByCmd("Brk #Tmp Split xxx Keep abc,def", ":") Then Stop
DoCmd.OpenTable "#Tmp_Brk"
Exit Function
E: Brk_Rec_ByCmd_Tst = True
End Function
#End If
Function Brk_Rec(ByVal pBrk$, ByVal pSplit$ _
        , Optional ByVal pInto$ = "" _
        , Optional ByVal pTo$ = "" _
        , Optional ByVal pKeep$ = "" _
        , Optional ByVal pSetSno$ = "" _
        , Optional ByVal pBeg% = 10 _
        , Optional ByVal pStp% = 10 _
        , Optional ByVal pSepChr$ = cComma _
        ) As Boolean
'Aim: Brk Split [Into] [To] [Keep] [Ord] [Beg] [Stp]
'           {pBrk}      each record
'           {pSplit}    split
'           {pTo}     into n records  use <pBrk>_Ln
'           {pInto}                       use Nm<pSplit>,   if <pSplit> begins with Ln
'                                       use <pSplit>_Brk  else
'           {pKeep}                     use <pBrk>
'           {pSetSno}
'           {pStep}
'Eg. pBrk = "Tbl", pSplit="LnFld"
'    Assume: Tbl     has fields: Tbl,LnFld
'    Then  : Tbl_Ln  has fields: Tbl,Sno,NmFld
'Eg. pBrk = "Tbl", pSplit="LoFld"
'    Assume: Tbl     has fields: Tbl,LoFld
'    Then  : Tbl_Ln  has fields: Tbl,Sno,LoFld_Brk
Const cSub$ = "Brk_Rec"
On Error GoTo R

'Set & Q with []: pBrk, pInto, pKeep, pTo, pOrd, pSplit
'Set            : pStp
If Left(pBrk, 1) = "(" And Right(pBrk, 1) = ")" Then
    If pInto = "" Then ss.A 1, "pBrk is (..), but pInto is not given": GoTo E
    If pKeep = "" Then ss.A 2, "pBrk is (..), but pKeep is not given": GoTo E
Else
    pBrk = jj.Rmv_SqBkt(pBrk)
    If pInto = "" Then pInto = pBrk & "_Brk"
    pBrk = jj.Q_SqBkt(pBrk)
End If
pInto = jj.Q_SqBkt(pInto)
'
If pKeep = "" Then
    pKeep = pBrk
Else
    If jj.Q_Ln(pKeep, pKeep, "[]") Then ss.A 3: GoTo E
End If
'
pSplit = jj.Rmv_SqBkt(pSplit)
If pTo = "" Then
    If Left(pSplit, 2) = "Ln" Then
        pTo = "Nm" & mID(pSplit, 3)
    Else
        pTo = pSplit & "_Ln"
    End If
End If
pSplit = jj.Q_SqBkt(pSplit)
pTo = jj.Q_SqBkt(pTo)
'
If pSetSno = "" Then pSetSno = "Sno" Else pSetSno = jj.Q_SqBkt(pSetSno)
'
If pBeg = 0 Then pBeg = 10
If pStp = 0 Then pStp = 10
'
'- Create {pInto}
Dim mSql$: mSql = jj.Fmt_Str("SELECT {0},CInt(0) as {1} into {2} from {3} where False" _
    , pKeep, pSetSno, pInto, pBrk)
If jj.Run_Sql(mSql) Then ss.A 1: GoTo E
mSql = jj.Fmt_Str("Alter Table {0} Add {1} Text(50)", pInto, pTo)
If jj.Run_Sql(mSql) Then ss.A 1: GoTo E

Dim mRsTar As DAO.Recordset, mRsSrc As DAO.Recordset, mAyKeep$(), NKeep%
Do
    mAyKeep = Split(pKeep, pSepChr):
    NKeep = jj.Siz_Ay(mAyKeep)
    Dim J%
    For J = 0 To NKeep - 1
        mAyKeep(J) = Trim(mAyKeep(J))
    Next
    If jj.Opn_Rs(mRsSrc, jj.Fmt_Str("Select {0},{1} from {2}", pKeep, pSplit, pBrk)) Then ss.A 1: GoTo E
    Set mRsTar = CurrentDb.TableDefs(pInto).OpenRecordset
Loop Until True
With mRsSrc
    While Not .EOF
        Dim mSno%: mSno = pBeg
        Dim mSplit$: mSplit = Trim(Nz(.Fields(pSplit).Value, ""))
        If mSplit <> "" Then
            Dim mAySplit$(): mAySplit = Split(mSplit, pSepChr)
            Dim xFm%, xTo%, xStep%
            If pStp > 0 Then
                xFm = 0: xTo = jj.Siz_Ay(mAySplit) - 1: xStep = 1
            Else
                xTo = 0: xFm = jj.Siz_Ay(mAySplit) - 1: xStep = -1
            End If
            Dim I%
            For I = xFm To xTo Step xStep
                mRsTar.AddNew
                For J = 0 To NKeep - 1
                    mRsTar.Fields(mAyKeep(J)).Value = .Fields(mAyKeep(J)).Value
                Next
                mRsTar.Fields(pSetSno).Value = mSno
                mRsTar.Fields(pTo).Value = Trim(mAySplit(I))
                mRsTar.Update
                mSno = mSno + pStp
            Next
        End If
        .MoveNext
    Wend
    .Close
End With
GoTo X
R: ss.R
E: Brk_Rec = True: ss.B cSub, cMod, "pBrk,pSplit,pInto,pTo,pKeep,pSetSno,pStp,pSepChr", pBrk, pSplit, pInto, pTo, pKeep, pSetSno, pStp, pSepChr
X: jj.Cls_Rs mRsTar, mRsSrc
End Function
#If Tst Then
Function Brk_Rec_Tst() As Boolean
Const cBrk$ = "#BrkRec"
Const cKeep = "abc,def"
Const cSplit = "xxx"
Const cInto = "xxx"
Dim mSql$:
'Create cNmtTo
Dim mCase As Byte: mCase = 4

Dim mBrk$, mSplit$, mInto$, mTo$, mKeep$, mOrd$, mStp%, mSepChr$
Select Case mCase
Case 1, 2
    'Create cBrk
    If jj.Dlt_Tbl(cBrk) Then Stop
    mSql = "Create table [" & cBrk & "] (abc text(10), def integer, xxx text(255))"
    If jj.Run_Sql(mSql) Then Stop
    Dim J%
    For J = 0 To 9
        Dim mA$: mA = "aa," & J & cComma & J + 1 & cComma & J + 2
        mSql = jj.Fmt_Str("Insert into [{0}] values ('abc-{1}',{2},'{3}')", cBrk, J, J * 2, mA)
        If jj.Run_Sql(mSql) Then Stop
    Next
    mBrk = cBrk
    mKeep = cKeep
    mSplit = cSplit
    mInto = cSplit
End Select

Select Case mCase
    'If qBrkRec(mBrk, mSplit,  mInto, mTo,mKeep, mOrd, mStp, mSepChr) Then Stop: GoTo E
Case 1
    If jj.Brk_Rec(mBrk, mSplit, mInto, mTo, mKeep, mOrd, mStp, mSepChr) Then Stop: GoTo E
    mTo = cBrk & "_Brk"
Case 2
    If jj.Brk_Rec(mBrk, mSplit, mInto, mTo, mKeep, mOrd, mStp, mSepChr) Then Stop: GoTo E
    mTo = cBrk & "_Brk"
Case 3
    mBrk = "#LgcT"
    mSplit = "LnFld"
    mInto = "#LgcTF"
    mTo = ""
    mKeep = "Lgc,NmLgcT"
    mOrd = "SnoLgcT"
    If jj.Brk_Rec(mBrk, mSplit, mInto, , mKeep, mOrd, mStp, mSepChr) Then Stop: GoTo E
Case 4
    mBrk$ = "#tmpTbl"
    mSplit = "LnFld"
    mInto = ""
    mTo = ""
    mKeep = "Tbl"
    mOrd = ""
    
    Dim cIns$: cIns = "Insert into [" & mBrk & "] values "
    If jj.Crt_Tbl_FmLoFld(mBrk, "Tbl Long, LnFld Text 255", 1) Then Stop: GoTo E
    If jj.Run_Sql(cIns & "(1,'aa,bb,dd,ee')") Then Stop: GoTo E
    If jj.Run_Sql("Insert into [#tmpTbl] values (2,'aa1,bb1,dd1')") Then Stop: GoTo E
    If jj.Run_Sql("Insert into [#tmpTbl] values (3,'x,y,z')") Then Stop: GoTo E
    If jj.Brk_Rec(mBrk, mSplit, , , mKeep) Then Stop: GoTo E
    mTo = mBrk & "_Brk"
End Select
DoCmd.OpenTable jj.Rmv_SqBkt(mBrk)
DoCmd.OpenTable jj.Rmv_SqBkt(mInto)
Exit Function
E:
    Brk_Rec_Tst = True
End Function
#End If
Function Brk_Cmb_RecCmd(oCmb$, oJoin$, oInto$, oTo$, oKeep$, oOrd$, oStp%, pCmbRecCmd$) As Boolean
'Aim: Break {pJnRecCmd} into: Cmb Jn [To] [Into] [Keep] [Ord] [Stp]
'     Assume no space with the elements
Const cSub$ = "Brk_Cmb_RecCmd"
Dim mCmbRecCmd$: mCmbRecCmd = Replace(Replace(Replace(pCmbRecCmd, vbLf, " "), vbCr, " "), "  ", " ")
Dim mA$(): mA = Split(mCmbRecCmd)
Dim J%
oCmb = "": oJoin = "": oTo = "": oInto = "": oKeep = "": oOrd = "": oStp = 0
For J = 0 To jj.Siz_Ay(mA) - 1 Step 2
    Select Case mA(J)
    Case "Cmb":     If oCmb <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oCmb = mA(J + 1)
    Case "Join":    If oJoin <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oJoin = mA(J + 1)
    Case "To":      If oTo <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oTo = mA(J + 1)
    Case "Into":    If oInto <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oInto = mA(J + 1)
    Case "Keep":    If oKeep <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oKeep = mA(J + 1)
    Case "Ord":     If oOrd <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oOrd = mA(J + 1)
    Case "Stp":     If oStp <> 0 Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oStp = Val(mA(J + 1))
    Case Else
                    ss.A 1, "Element(" & mA(J) & ") is not expected is more than one": GoTo E
    End Select
Next
Exit Function
E: Brk_Cmb_RecCmd = True: ss.B cSub, cMod, "pCmbRecCmd", pCmbRecCmd
End Function
Function Brk_Brk_RecCmd(oBrk$, oSplit$, oInto$, oTo$, oKeep$, oSetSno$, oBeg%, oStp%, pBrkRecCmd$) As Boolean
'Aim: Break {pBrkRecCmd} into: Brk Split [To] [Into] [Keep] [SetSno] [Beg] [Stp]
'     Assume no space with the elements
Const cSub$ = "Brk_Brk_RecCmd"
Dim mBrkRecCmd$: mBrkRecCmd = Replace(Replace(Replace(pBrkRecCmd, vbLf, " "), vbCr, " "), "  ", " ")
Dim mA$(): mA = Split(mBrkRecCmd)
Dim J%
oBrk = "": oSplit = "": oTo = "": oInto = "": oKeep = "": oSetSno = "": oStp = 0
For J = 0 To jj.Siz_Ay(mA) - 1 Step 2
    Select Case mA(J)
    Case "Brk":     If oBrk <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oBrk = mA(J + 1)
    Case "Split":   If oSplit <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oSplit = mA(J + 1)
    Case "To":      If oTo <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oTo = mA(J + 1)
    Case "Into":    If oInto <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oInto = mA(J + 1)
    Case "Keep":    If oKeep <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oKeep = mA(J + 1)
    Case "SetSno":  If oSetSno <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oSetSno = mA(J + 1)
    Case "Beg":     If oBeg <> 0 Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oBeg = Val(mA(J + 1))
    Case "Stp":     If oStp <> 0 Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oStp = Val(mA(J + 1))
    Case Else
                    ss.A 1, "Element(" & mA(J) & ") is not expected.", , "Expected Values", "Brk Split Into To Keep SetSno Stp": GoTo E
    End Select
Next
Exit Function
E: Brk_Brk_RecCmd = True: ss.B cSub, cMod, "pBrkRecCmd", pBrkRecCmd
End Function
Function Brk_ClsBracket(o1$, o2$, pS$) As Boolean
'Aim: Break pS$ into o1 & o2 by a ')'
Const cSub$ = "Brk_ClsBracket"
Dim mA$: mA = Replace(pS, "()", "  ")
Dim mP%: mP = InStr(mA, ")")
If mP = 0 Then o1 = pS: o2 = "": Exit Function
o1 = Left(pS, mP - 1)
o2 = mID(pS, mP + 1)
End Function
#If Tst Then
Function Brk_ClsBracket_Tst() As Boolean
Dim mS1$, mS2$: If jj.Brk_ClsBracket(mS1, mS2, "dsfjd()lskdfj)lskdjf") Then Stop
Debug.Print mS1
Debug.Print mS2
End Function
#End If
Function Brk_PrcBody(oDPgm As d_Pgm, oAyDArg() As d_Arg, pPrcBody$) As Boolean
'Aim: Assume all remark lines before "Const cSub$" is "Aim"
Const cSub$ = "Brk_PrcBody"
jj.Clr_AyDArg oAyDArg
jj.Clr_DPgm oDPgm
Dim mPrcBody$: mPrcBody = Trim(pPrcBody)
Dim mP%, mA$, mB$
Dim mToken As eSts: mToken = ePub_Prv
mToken = eSts.ePub_Prv
With oDPgm
    .x_PrcBody = pPrcBody
    Do Until mToken = EEnd
        Select Case mToken
        Case ePub_Prv
            mP = InStr(mPrcBody, " "): mA = Left(mPrcBody, mP): mPrcBody = mID(mPrcBody, mP + 1)
            Select Case mA
            Case "Private ":    mToken = ePrp_Fct_Sub:  .x_IsPrivate = True
            Case "Public ":     mToken = ePrp_Fct_Sub
            Case "Property ":   mToken = eGet_Set_Let
            Case "Function Brk_":   mToken = eNmPrc:    .x_TypFct = eFct
            Case "Sub ":        mToken = eNmPrc:    .x_TypFct = eSub
            Case Else:          ss.A 1, "Unexpected token", , "Cur Token,Expected Token", mA, "Private,Public,Property,Function,Sub": GoTo E
            End Select
        Case ePrp_Fct_Sub
            mP = InStr(mPrcBody, " "): mA = Left(mPrcBody, mP): mPrcBody = mID(mPrcBody, mP + 1)
            Select Case mA
            Case "Property ":   mToken = eGet_Set_Let
            Case "Function Brk_":   mToken = eNmPrc:    .x_TypFct = eSub
            Case "Sub ":        mToken = eNmPrc:    .x_TypFct = eSub
            Case Else:          ss.A 1, "Unexpected token", , "Cur Token,Expected Token", mA, "Property,Function,Sub": GoTo E
            End Select
        Case eGet_Set_Let
            mP = InStr(mPrcBody, " "): mA = Left(mPrcBody, mP): mPrcBody = mID(mPrcBody, mP + 1)
            Select Case mA
            Case "Get ":    mToken = eNmPrc:    .x_TypFct = eGet
            Case "Set ":    mToken = eNmPrc:    .x_TypFct = eSet
            Case "Let ":    mToken = eNmPrc:    .x_TypFct = eLet
            Case Else:      ss.A 1, "Unexpected token", , "Cur Token,Expected Token", mA, "Get,Set,Let": GoTo E
            End Select
        Case eNmPrc
            mP = InStr(mPrcBody, "("): mA = Left(mPrcBody, mP - 1): mPrcBody = mID(mPrcBody, mP + 1)
            mB = Right(mA, 1)
            Select Case mB
            Case "$", "&", "%", "#", "!", "@":  .x_NmTypRet = mB: mA = Left(mA, Len(mA) - 1)
            Case Else:                          .x_NmTypRet = ""
            End Select
            .x_NmPrc = mA
            mToken = eArg
        Case eArg
            'Assume there is no ' in arguement list.
            If jj.Brk_ClsBracket(mA, mB, mPrcBody) Then ss.A 1: GoTo E
            'Note: for        xxx(...) AS xxx
            '      mPrcBody =     ...) AS xxx
            '      mA       =     ...
            '      mB       =      AS xxx
            'Note: mPrcBody = 'Aim...
            If jj.Brk_PrmDcl(oAyDArg, mA) Then ss.A 3: GoTo E
            mPrcBody = Trim(mB) ' As .....
            If Left(mPrcBody, 3) = "As " Then
                mA = mID(mPrcBody, 4): mP = InStr(mA, vbCrLf): If mP > 0 Then mA = Left(mA, mP - 1)
                mP = InStr(mA, " ")
                If mP > 0 Then mA = Left(mA, mP - 1)
                If Right(mA, 1) = ":" Then
                    .x_NmTypRet = Left(mA, Len(mA) - 1)
                Else
                    .x_NmTypRet = mA
                End If
                Select Case .x_NmTypRet
                Case "String": .x_NmTypRet = "$"
                Case "Currency": .x_NmTypRet = "@"
                Case "Integer": .x_NmTypRet = "%"
                Case "Double": .x_NmTypRet = "#"
                Case "Single": .x_NmTypRet = "!"
                Case "Long": .x_NmTypRet = "&"
                End Select
            End If
            
            mToken = eAim
        Case eAim
            mP = InStr(mPrcBody, "'Aim")
            If mP > 0 Then
                .x_Aim = jj.Cut_NonRmk(mID(mPrcBody, mP))
            End If
            Exit Function
        End Select
    Loop
End With
Exit Function
E: Brk_PrcBody = True: ss.B cSub, cMod, "pPrcBody", pPrcBody
End Function
#If Tst Then
Function Brk_PrcBody_Tst() As Boolean
Const cSub$ = "Brk_PrcBody_Tst"
Dim mFb$:
Dim mDPgm As d_Pgm, mAyDArg() As d_Arg, mPrcBody$
Dim mNmPrc$, mNmPrj_Nmm$

Dim mCase As Byte: mCase = 7
Select Case mCase
Case 1: mNmPrj_Nmm = "JMtcDb.RunGenMdb":    mNmPrc = "qryGenMdb_CrtMdb_Fm_Mdb_Run":     mFb = "p:\workingdir\pgmobj\JMtcDb.Mdb"
Case 2: mNmPrj_Nmm = "jj.Add":              mNmPrc = "Str":                             mFb = ""
Case 3: mNmPrj_Nmm = "jj.Brk":              mNmPrc = "PrmDcl":                          mFb = ""
Case 4: mNmPrj_Nmm = "jj.Add":              mNmPrc = "Ay":                              mFb = ""
Case 5: mNmPrj_Nmm = "jj.Bld":              mNmPrc = "Lst_ByAyV":                       mFb = ""
Case 6: mNmPrj_Nmm = "jj.Dlt":              mNmPrc = "Tbl":                             mFb = ""
Case 7: mNmPrj_Nmm = "jj.Fmt":              mNmPrc = "yMmmWww":                         mFb = ""
End Select

Dim mAcs As Access.Application: If jj.Cv_Acs_FmFb(mAcs, mFb) Then Stop: GoTo E
If jj.Fnd_PrcBody(mPrcBody, mNmPrj_Nmm, mNmPrc, mAcs) Then Stop: GoTo E
If jj.Brk_PrcBody(mDPgm, mAyDArg, mPrcBody$) Then Stop: GoTo E
jj.Shw_Dbg cSub, cMod, "mDPgm,mAyDArg,mPrcBody", jj.ToStr_DPgm(mDPgm), jj.ToStr_AyDArg(mAyDArg), mPrcBody
Exit Function
E: Brk_PrcBody_Tst = True
X:
    If mFb <> "" Then jj.Cls_CurDb mAcs
End Function
#End If
Function Brk_PrmDcl(oAyDArg() As d_Arg, pPrmDcl$) As Boolean
'Aim: Brk pPrcBody in fmt (...) into oAyDArg()
Const cSub$ = "BrkArg"
If pPrmDcl = "" Then jj.Clr_AyDArg oAyDArg: Exit Function
Dim mAyArgDcl$(): mAyArgDcl = Split(Replace(Replace(pPrmDcl, "_" & vbCrLf, " "), vbCrLf, " "), cComma)
Dim J%, N%: N% = jj.Siz_Ay(mAyArgDcl)
ReDim oAyDArg(N - 1)
For J = 0 To N - 1
    If jj.Brk_ArgDcl(oAyDArg(J), mAyArgDcl(J)) Then ss.A 1: GoTo E
Next
Exit Function
E: Brk_PrmDcl = True: ss.B cSub, cMod, "pPrmDcl", pPrmDcl
End Function
#If Tst Then
Function Brk_PrmDcl_Tst() As Boolean
Dim mPrmDcl$, mAyDArg() As d_Arg
Dim mPrcBody$, mNmPrc$, mNmPrj_Nmm$
Dim mCase As Byte
mCase = 7
Select Case mCase
Case 1: mPrmDcl = "oDArg As d_Arg, pArgDcl$"
Case 2: mPrmDcl = "Optional pArgDcl As String = ""ABC"""
Case 3: mPrmDcl = "Optional pInclTbl As Boolean = True, Optional ByVal pInclQry As Boolean = True, Optional pInclTypFld As Boolean = False, Optional pCls As Boolean = False"
Case 4: mNmPrj_Nmm = "jj.Bld":      mNmPrc = "Lst_ByAyV":
Case 5: mNmPrj_Nmm = "jj.ToStr":    mNmPrc = "Ays"
Case 6: mNmPrj_Nmm = "jj.ToStr":    mNmPrc = "Ays"
Case 7: mNmPrj_Nmm = "jj.Run":      mNmPrc = "qBrkRec"
End Select
If 4 <= mCase And mCase <= 7 Then
    If jj.Fnd_PrcBody(mPrcBody, mNmPrj_Nmm, mNmPrc) Then Stop: GoTo E
    mPrmDcl = jj.Cut_Prm(mPrcBody)
End If
If jj.Brk_PrmDcl(mAyDArg, mPrmDcl) Then Stop: GoTo E
jj.Shw_DbgWin

Debug.Print jj.Fct.UnderlineStr(mPrmDcl, "*")
Debug.Print mPrmDcl
Debug.Print jj.Fct.UnderlineStr(mPrmDcl)
Dim J%
For J = 0 To jj.Siz_AyDArg(mAyDArg) - 1
    Debug.Print jj.ToStr_DArg(mAyDArg(J))
Next
Exit Function
E: Brk_PrmDcl_Tst = True
End Function
#End If
Function Brk_ArgDcl(oDArg As d_Arg, pArgDcl$) As Boolean
'Aim: Brk pPrcBody in fmt (...) into oAyDArg()
Const cSub$ = "BrkArgDcl"
'    Public x_IsAs Boolean, x_NmArg$, x_NmTypArg$, x_IsOpt As Boolean, x_DftVal
If TypeName(oDArg) = "Nothing" Then Set oDArg = New d_Arg
Dim mArgDcl$: mArgDcl = Trim(Replace(pArgDcl, "_" & vbCrLf, ""))
With oDArg
    Dim mP%, mA$
    .x_IsOpt = False
    .x_IsByVal = False
    .x_IsAy = False
    .x_IsPrmAy = False
    mA = "Optional ":   mP = InStr(mArgDcl, mA): .x_IsOpt = (mP > 0):    If mP > 0 Then mArgDcl = Trim(Replace(mArgDcl, mA, ""))
    mA = "ByVal ":      mP = InStr(mArgDcl, mA): .x_IsByVal = (mP > 0):  If mP > 0 Then mArgDcl = Trim(Replace(mArgDcl, mA, ""))
    mA = "ByRef ":      mP = InStr(mArgDcl, mA):                         If mP > 0 Then mArgDcl = Trim(Replace(mArgDcl, mA, ""))
    mA = "ParamArray ": mP = InStr(mArgDcl, mA): .x_IsPrmAy = (mP > 0):  If mP > 0 Then mArgDcl = Trim(Replace(mArgDcl, mA, ""))
    mA = "()":          mP = InStr(mArgDcl, mA): .x_IsAy = (mP > 0):     If mP > 0 Then mArgDcl = Trim(Replace(mArgDcl, mA, ""))
    
    mP = InStr(mArgDcl, " = ")
    If mP > 0 Then
        .x_DftVal = Trim(mID(mArgDcl, mP + 3))
        mArgDcl = Trim(Left(mArgDcl, mP - 1))
    Else
        .x_DftVal = ""
    End If
    
    mP = InStr(mArgDcl, " As ")
    If mP > 0 Then
        .x_NmTypArg = Trim(mID(mArgDcl, mP + 3))
        Select Case .x_NmTypArg
        Case "String": .x_NmTypArg = "$"
        Case "Long": .x_NmTypArg = "&"
        Case "Integer": .x_NmTypArg = "%"
        Case "Single": .x_NmTypArg = "!"
        Case "Double": .x_NmTypArg = "#"
        Case "Currency": .x_NmTypArg = "@"
        End Select
        .x_NmArg = Trim(Left(mArgDcl, mP - 1))
    Else
        Dim mB$
        mB = Right(mArgDcl, 1)
        Select Case mB
        Case "%", "$", "&", "#", "!":   .x_NmTypArg = mB:           .x_NmArg = Left(mArgDcl, Len(mArgDcl) - 1)
        Case Else:                      .x_NmTypArg = "Variant":    .x_NmArg = mArgDcl
        End Select
    End If
End With
Exit Function
E: Brk_ArgDcl = True: ss.B cSub, cMod, "pArgDcl", pArgDcl
End Function
#If Tst Then
Function Brk_ArgDcl_Tst() As Boolean
Dim mArgDcl$
Dim mDArg As New d_Arg
Dim mCase As Byte
mCase = 2
Select Case mCase
Case 1: mArgDcl = "ByVal pArgDcl$"
Case 2: mArgDcl = "Optional ByVal pArgDcl As String = ""ABC"""
Case 3: mArgDcl = "pArgDcl$()"
End Select
If jj.Brk_ArgDcl(mDArg, mArgDcl) Then Stop: GoTo E
jj.Shw_DbgWin

Debug.Print jj.Fct.UnderlineStr(mArgDcl, "*")
Debug.Print mArgDcl
Debug.Print jj.Fct.UnderlineStr(mArgDcl)

Debug.Print jj.ToStr_DArg(mDArg, False)
Exit Function
E: Brk_ArgDcl_Tst = True
End Function
#End If
Function Brk_Sql_ToAnKw(oAnKw$(), oAyTypKw() As Byte, oAyKwLen() As Byte, pSql$) As Boolean
'Aim: break {pSql} to {oAnKw} with {oAyTypKw} & {oAyKwLen}.  TypKw:1=From;2=Into;9=Other
Const cSub$ = "Brk_Sql_ToAnKw"
jj.Clr_Ays oAnKw
jj.Clr_AyByt oAyTypKw
jj.Clr_AyByt oAyKwLen
Dim mSql$: mSql = RTrim(Replace(Replace(pSql, vbLf, " "), vbCr, " "))
Dim p%, mTypKw As Byte, mKwLen As Byte: p = InStr_SqlKw%(mTypKw, mKwLen, mSql)
While p > 0
    Dim mKw$: mKw = Right(mSql, Len(mSql) - p + 1)
    jj.Add_AyEle oAnKw, mKw
    jj.Add_AyByt oAyKwLen, mKwLen
    jj.Add_AyByt oAyTypKw, mTypKw
    mSql = Left(mSql, p - 1)
    p = InStr_SqlKw%(mTypKw, mKwLen, mSql)
Wend
End Function
Function Brk_Lnt(oAnt$(), pLnt$) As Boolean
jj.Clr_Ays oAnt
oAnt = Split(pLnt, cComma)
Dim J%
For J = 0 To jj.Siz_Ay(oAnt) - 1
    oAnt(J) = jj.Rmv_SqBkt(jj.Fct.RTrimSemiQ(oAnt(J)))
Next
End Function
Function Brk_Sql_ToAnt(oAnt$(), pSql$) As Boolean
Const cSub$ = "Brk_Sql_ToAnt"
Dim mAnKw$(), mAyTypKw() As Byte, mAyKwLen() As Byte
If jj.Brk_Sql_ToAnKw(mAnKw, mAyTypKw, mAyKwLen, pSql) Then Stop: GoTo E
Dim J%
jj.Clr_Ays oAnt
For J = 0 To jj.Siz_Ay(mAyTypKw) - 1
    If mAyTypKw(J) = 1 Then
        Dim mAnt$(): If jj.Brk_Lnt(mAnt, mID(mAnKw(J), mAyKwLen(J))) Then Stop: GoTo E
        jj.Add_AyAtEnd oAnt, mAnt
    End If
Next
Exit Function
E: Brk_Sql_ToAnt = True
End Function
#If Tst Then
Function Brk_Sql_ToAnt_Tst() As Boolean
Const cFt$ = "c:\aa.csv"
If jj.Dlt_Fil(cFt) Then Stop: GoTo E
Dim mF As Byte: mF = FreeFile: Open cFt For Output As #mF
Print #mF, "Nmq,Lnt,Sql"
Dim mAnq$(), mAnt$(): If jj.Fnd_Anq_ByLik(mAnq, "q*") Then Stop: GoTo E
Dim J%
For J = 0 To jj.Siz_Ay(mAnq) - 1
    Debug.Print mAnq(J)
    Dim mSql$: mSql = CurrentDb.QueryDefs(mAnq(J)).Sql
    If jj.Brk_Sql_ToAnt(mAnt, mSql) Then Stop: GoTo E
    Dim I%
    For I = 0 To jj.Siz_Ay(mAnt) - 1
        Write #mF, mAnq(J), jj.ToStr_Ays(mAnt), mSql
    Next
Next
Close #mF
Dim mWb As Workbook: If jj.Opn_Wb(mWb, cFt, , , True) Then Stop
Exit Function
E: Brk_Sql_ToAnt_Tst = True
End Function
#End If
#If Tst Then
Function Brk_Sql_ToAnKw_Tst() As Boolean
Const cFt$ = "c:\aa.csv"
If jj.Dlt_Fil(cFt) Then Stop: GoTo E

Dim mF As Byte: mF = FreeFile: Open cFt For Output As #mF
Print #mF, "Nmq,TypKw,Kw,Lnt,CleanLnt,Sql"

Dim mAnq$(): If jj.Fnd_Anq_ByLik(mAnq, "qry*") Then Stop: GoTo E

Dim J%
For J = 0 To jj.Siz_Ay(mAnq) - 1
    Debug.Print mAnq(J)
    Dim mSql$, mAnKw$(), mAyTypKw() As Byte, mAyKwLen() As Byte
    mSql = CurrentDb.QueryDefs(mAnq(J)).Sql
    If jj.Brk_Sql_ToAnKw(mAnKw, mAyTypKw, mAyKwLen, mSql) Then Stop: GoTo E
    Dim I%
    For I = 0 To jj.Siz_Ay(mAnKw) - 1
        Dim mLnt$: mLnt = mID(mAnKw(I), mAyKwLen(I))
        Dim mAnt$(): If jj.Brk_Lnt(mAnt, mLnt) Then Stop: GoTo E
        Write #mF, mAnq(J), mAyTypKw(I), mAnKw(I), mLnt, jj.ToStr_Ays(mAnt), mSql
    Next
Next
Close #mF
Dim mWb As Workbook: If jj.Opn_Wb(mWb, cFt, , , True) Then Stop
Exit Function
E: Brk_Sql_ToAnKw_Tst = True
End Function
#End If
Function Brk_Nm_InTbl(pNmt$, pNmFld$, Optional pMax As Byte = 5) As Boolean
'Aim: Break the Nm into Nm1,..Nm5
Const cSub$ = "Brk_Nm_InTbl"
With CurrentDb.TableDefs(pNmt).OpenRecordset
    While Not .EOF
        .Edit
        Dim mAy$(): If jj.Brk_Nm(mAy, .Fields(pNmFld).Value, pMax) Then ss.A 1: GoTo E
        .Fields(pNmFld & "1").Value = mAy(0)
        .Fields(pNmFld & "2").Value = mAy(1)
        .Fields(pNmFld & "3").Value = mAy(2)
        .Fields(pNmFld & "4").Value = mAy(3)
        .Fields(pNmFld & "5").Value = mAy(4)
        .Update
        .MoveNext
    Wend
    .Close
End With
Exit Function
E: Brk_Nm_InTbl = True
End Function
Function Brk_Nm(oAy$(), pNm$, Optional pMax As Byte = 5) As Boolean
'Aim: Break Nm into Nm1,..,Nm5
Const cSub$ = "Brk_Nm"
Const cA As Byte = 65
Const cZ As Byte = 90
Dim J%, mS As Byte, mA$: mA = ""
jj.Clr_Ays oAy
For J = 1 To Len(pNm)
    mS = Asc(mID(pNm, J, 1))
    If cA <= mS And mS <= cZ Then
        If Len(mA) > 0 Then
            If jj.Add_AyEle(oAy, mA) Then ss.A 1: GoTo E
            mA = ""
        End If
    End If
    mA = mA & Chr(mS)
Next
If Len(mA) > 0 Then If jj.Add_AyEle(oAy, mA) Then ss.A 2: GoTo E
Dim N%: N = jj.Siz_Ay(oAy)
If N > pMax Then
    mA = ""
    For J = pMax - 1 To N - 1
        mA = mA & oAy(J)
    Next
    oAy(pMax - 1) = mA
End If
ReDim Preserve oAy(pMax - 1)
Exit Function
E: Brk_Nm = True: ss.B cSub, cMod, "pNm", pNm
End Function
#If Tst Then
Function Brk_Nm_Tst() As Boolean
Dim mA$: mA = "A1A2A3A4A5A6A7"
Dim mAy$(): If jj.Brk_Nm(mAy, mA) Then Stop
Debug.Print mA
Debug.Print jj.UnderlineStr(mA)
Debug.Print jj.ToStr_Ays(mAy, , vbLf)
End Function
#End If
Function Brk_ColonAs_ToCaptionNm(oCaption$, oNm$, pColonAsStr$) As Boolean
'Aim: Convert "[<<Caption>>:] <<Nam>>" into oNm and oCaption
Const cSub$ = "Brk_ColonAs_ToCaptionNm"
If jj.Brk_Str_1ForS2(oCaption, oNm, pColonAsStr, ":") Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: Brk_ColonAs_ToCaptionNm = True: ss.B cSub, cMod, "pColonAsStr", pColonAsStr
End Function
#If Tst Then
Function Brk_ColonAs_ToCaptionNm_Tst() As Boolean
Const cSub$ = "Brk_ColonAs_ToCaptionNm_Tst"
jj.Shw_Dbg cSub, cMod
Dim mColonAsStr$, mNm$, mCaption$, mCase As Byte
For mCase = 1 To 2
    Select Case mCase
    Case 1
        mColonAsStr = "aa: xx"
    Case 2
        mColonAsStr = "xx"
    End Select
    If jj.Brk_ColonAs_ToCaptionNm(mCaption, mNm, mColonAsStr) Then Stop
    Debug.Print mCase
    Debug.Print jj.ToStr_LpAp(vbLf, "mColonAsStr, mNm, mCaption", mColonAsStr, mNm, mCaption)
    Debug.Print "-----------"
Next
End Function
#End If
Function Brk_Lin2Ays(oAys$(), pLin$) As Boolean
oAys = Split(pLin, vbCrLf)
Dim J%
For J = 0 To UBound(oAys)
    oAys(J) = Replace(oAys(J), "|", vbCrLf)
Next
End Function
Function Brk_LpVv2Am(oAm() As tMap, pLp$, pVayv) As Boolean
Dim mA$(): mA = Split(pLp, cComma)
Dim mAy(): mAy = pVayv
Dim mN1%: mN1 = jj.Siz_Ay(mA)
Dim mN2%: mN2 = jj.Siz_Vayv(pVayv)
Dim mN%: mN = jj.Fct.MaxInt(mN1, mN2)
If mN1 = 0 And mN2 = 0 Then jj.Clr_Am oAm: Exit Function
ReDim oAm(mN) As tMap
Dim J%
For J = 0 To mN - 1
    If J < mN1 Then oAm(J).F1 = mA(J)
    If J < mN2 Then If Not IsMissing(mAy(J)) Then oAm(J).F2 = Nz(mAy(J), "")
Next
End Function
Function Brk_LpAp2Am(oAm() As tMap, pLp$, ParamArray pAp()) As Boolean
Brk_LpAp2Am = jj.Brk_LpVv2Am(oAm, pLp, CVar(pAp))
End Function
Function Brk_LpAp2Am_Tst() As Boolean
Dim mAm() As tMap, J%
If jj.Brk_LpAp2Am(mAm, "string,num,date", "12", 34, Date) Then Stop
For J = 0 To jj.Siz_Am(mAm) - 1
    Debug.Print jj.ToStr_Map(mAm(J))
Next
End Function
Function Brk_Lv2AyV(pLv$, Optional pSepChr$ = cComma) As Variant()
Dim mAys$(): mAys = Split(pLv, pSepChr)
Dim N%, J%: N = jj.Siz_Ay(mAys): If N = 0 Then Exit Function
ReDim mAyV(N - 1)
For J = 0 To N - 1
    mAyV(J) = mAys(J)
Next
Brk_Lv2AyV = mAyV
End Function
Function Brk_Lv(pLv$, oV0$, oV1$ _
    , Optional oV2$ _
    , Optional oV3$ _
    , Optional oV4$ _
    , Optional oV5$ _
    , Optional oV6$ _
    , Optional oV7$ _
    , Optional oV8$ _
    , Optional oV9$ _
    , Optional oV10$ _
    , Optional oV11$ _
    , Optional oV12$ _
    , Optional oV13$ _
    , Optional oV14$ _
    , Optional oV15$) As Boolean
Const cSub$ = "Brk_Prm"
'Aim:  Break pLv$ into oV0, ...
'Note: Each value in pLv is separated by vbCrLf.  After Break, any | in each value will be replaced by vbCrLf
On Error GoTo R
Dim mAy$(): mAy = Split(pLv, vbCrLf)
Dim J%
For J = 0 To UBound(mAy)
    Select Case J
        Case 0: oV0 = Replace(mAy(0), "|", vbCrLf)
        Case 1: oV1 = Replace(mAy(1), "|", vbCrLf)
        Case 2: oV2 = Replace(mAy(2), "|", vbCrLf)
        Case 3: oV3 = Replace(mAy(3), "|", vbCrLf)
        Case 4: oV4 = Replace(mAy(4), "|", vbCrLf)
        Case 5: oV5 = Replace(mAy(5), "|", vbCrLf)
        Case 6: oV6 = Replace(mAy(6), "|", vbCrLf)
        Case 7: oV7 = Replace(mAy(7), "|", vbCrLf)
        Case 8: oV8 = Replace(mAy(8), "|", vbCrLf)
        Case 9: oV9 = Replace(mAy(9), "|", vbCrLf)
        Case 10: oV10 = Replace(mAy(10), "|", vbCrLf)
        Case 11: oV11 = Replace(mAy(11), "|", vbCrLf)
        Case 12: oV12 = Replace(mAy(12), "|", vbCrLf)
        Case 13: oV13 = Replace(mAy(13), "|", vbCrLf)
        Case 14: oV14 = Replace(mAy(14), "|", vbCrLf)
        Case 15: oV15 = Replace(mAy(15), "|", vbCrLf)
    End Select
Next
Exit Function
R: ss.R
E: Brk_Lv = True: ss.B cSub, cMod, "pLv", pLv
End Function
#If Tst Then
Function Brk_Lv_Tst() As Boolean
Const cSub$ = "Brk_Lv_Tst"
Dim mV0$, mV1$, mV2$, mV3$, mV4$, mV5$, mV6$, mV7$, mV8$, mV9$, mV10$
Dim mA$: mA = jj.Join_Prm("lkdsf|sfsdf", "lksdjf", "lskdjf")
If jj.Cut_Lv(mA, mV0, mV1, mV2) Then Stop
jj.Shw_Dbg cSub, cMod, "mA,0,1,2", mA, mV0, mV1, mV2
End Function
#End If
Function Brk_LoKwv(oAyKw$(), oAyKv$(), pLoKwv$, Optional pBrkChr$ = cComma) As Boolean
Const cSub$ = "Brk_LoKwv"
'Aim: Break {pLoKwv} into {oAyKw} & {oAyKv}
'     pLoKwv Fmt: Kw(Kv), Kw(Kv)
Dim mAy$(): mAy = Split(pLoKwv, pBrkChr)
Dim J%, I%, mA$, mKw$, mKv$
I = 0
For J = 0 To UBound(mAy)
    mA = Trim(mAy(J))
    If mA <> "" Then
        If jj.Brk_Kwv(mKw, mKv, mA) Then ss.A 1: GoTo E
        ReDim Preserve oAyKw(I), oAyKv(I)
        oAyKw(I) = mKw
        oAyKv(I) = mKv
        I = I + 1
    End If
Next
Exit Function
R: ss.R
E: Brk_LoKwv = True: ss.B cSub, cMod, "pLoKwv,pBrkChr", pLoKwv, pBrkChr
End Function
#If Tst Then
Function Brk_LoKwv_Tst() As Boolean
Const cSub$ = "Brk_LoKwv_Tst"
Dim mAyKw$(), mAyKv$()
Dim mA$: mA = "adfd(dfdf),dlfj(dfdffd)"
If jj.Brk_LoKwv(mAyKw, mAyKv, mA) Then Stop
jj.Shw_Dbg cSub, cMod
Debug.Print mA
Debug.Print "--"
Debug.Print jj.ToStr_Ays(mAyKw)
Debug.Print "--"
Debug.Print jj.ToStr_Ays(mAyKv)
End Function
#End If
Function Brk_Kwv(oKw$, oKv$, pKwv$) As Boolean
'Aim: Brk {pKwv} (fmt: kw(kv))
Const cSub$ = "Brk_Kwv"
Dim mA$: mA = Trim(pKwv)
Dim p1 As Byte, p2 As Byte
p1 = InStr(mA, "("): If p1 <= 0 Then ss.A 1, "pKwv has no [(]": GoTo E
p2 = InStr(p1 + 1, mA, ")"): If p2 <= 0 Then ss.A 2, "pKwv has no [)] after [(]": GoTo E
oKw = Left(mA, p1 - 1)
oKv = mID(mA, p1 + 1, p2 - p1 - 1)
Exit Function
E: Brk_Kwv = True: ss.B cSub, cMod, "pKwv", pKwv
End Function
Function Brk_Lm_To2Ay(oAy1$(), oAy2$(), pLm$ _
    , Optional pBrkChr$ = "=" _
    , Optional pSepChr$ = cComma) As Boolean
'Aim: Brk pLm (Fmt: aaa=xxx,bbb=yyy,1111) into oAy1$ (aaa,bbb,1111) and oAy2$ (xxx,bbb,1111)
oAy1 = Split(pLm, pSepChr)
Dim N%: N = jj.Siz_Ay(oAy1)
ReDim oAy2(N - 1)
Dim J%, p%, L%
L = Len(pBrkChr)
For J = 0 To N - 1
    oAy1(J) = Trim(oAy1(J))
    p = InStr(oAy1(J), pBrkChr)
    Select Case p
    Case 0: oAy2(J) = oAy1(J)
    Case 1: oAy1(J) = mID(oAy1(J), 1 + L): oAy2(J) = oAy1(J)
    Case Else: oAy2(J) = mID(oAy1(J), p + L): oAy1(J) = Left(oAy1(J), p - 1)
    End Select
Next
End Function
#If Tst Then
Function Brk_Lm_To2Ay_Tst() As Boolean
Const cSub$ = "Brk_Lm_To2Ay_Tst"
Dim mLm$, mSepChr$, mBrkChr$, mCase As Byte
Dim mAy1$(), mAy2$()
jj.Shw_Dbg cSub, cMod
For mCase = 1 To 2
    Select Case mCase
    Case 1
        mLm = "aaa=xxx,bbb=yyy,1111"
        mBrkChr = "="
        mSepChr = cComma
    Case 2
        mLm = "aaa as ccc, dd, ee as xx"
        mBrkChr = " as "
        mSepChr = cComma
    End Select
    If jj.Brk_Lm_To2Ay(mAy1, mAy2, mLm, mBrkChr, mSepChr) Then Stop
    Debug.Print "Input----"
    Debug.Print "mLm="; mLm
    Debug.Print "Output----"
    Debug.Print "mAy1="; jj.ToStr_Ays(mAy1)
    Debug.Print "mAy2="; jj.ToStr_Ays(mAy2)
Next
End Function
#End If
Function Brk_Ffn_To2Seg(oFfnn$, oExt$, pFfn$) As Boolean
Dim mPosDot As Byte: mPosDot = InStrRev(pFfn, ".")
If mPosDot = 0 Then oFfnn = pFfn: oExt = "": Exit Function
oFfnn = Left(pFfn, mPosDot - 1)
oExt = mID(pFfn, mPosDot)
End Function
Function Brk_Ffn_To3Seg(oDir$, oFnn$, oExt$, pFfn$) As Boolean
oDir = Fct.Nam_DirNam(pFfn)
Brk_Ffn_To3Seg = jj.Brk_Ffn_To2Seg(oFnn, oExt, Fct.Nam_FilNam(pFfn))
End Function
Function Brk_Ln2Ay(oAy$(), pLn$, Optional pSepChr$ = cComma) As Boolean
'Aim: Break list of string in {pLn} into {oAy} with separated by {pSepChr}.  Each element in {oAy} will always be trimed
Const cSub$ = "Brk_Ln2Ay"
On Error GoTo R
oAy = Split(pLn, pSepChr)
Dim J%: For J = 0 To jj.Siz_Ay(oAy) - 1
    oAy(J) = Trim(oAy(J))
Next
Exit Function
R: ss.R
E: Brk_Ln2Ay = True: ss.B cSub, cMod, "pLn,pSepChr", pLn, pSepChr$
End Function
Function Brk_Q(pQ$, oQ1$, oQ2$) As Boolean
Const cSub$ = "Brk_Q"
Select Case Len(pQ)
Case 0: oQ1 = "": oQ2 = "": Exit Function
Case 1: oQ1 = pQ: oQ2 = pQ: Exit Function
Case 2: oQ1 = Left(pQ, 1): oQ2 = Right(pQ, 1)
Case Else
    If jj.Brk_Str_1Or2(oQ1, oQ2, pQ, "*", True) Then ss.A 1, "quote string pQ is longer than 2, but no * inside.  It returns *blank quote": GoTo E
End Select
Exit Function
E: Brk_Q = True: ss.B cSub, cMod, "pQ", pQ
End Function
Function Brk_Str_Both(oS1$, oS2$, pS$, Optional pBrkChr$ = "=", Optional pNoTrim As Boolean = False) As Boolean
'Aim: Brk {pS} into {oS1} & {oS2}.  Format of pS: <oS1><pBrkChr><oS2>) with both <oS1> & <oS2> & <pBrkChr> must exist
Const cSub$ = "Brk_Str_Both"
If pBrkChr = "" Then ss.A 1, "BrkChr must be given": GoTo E
Dim p%: p = InStr(pS, pBrkChr): If p = 0 Then ss.A 2, "pS must contain pBrkChr": GoTo E
If pNoTrim Then
    oS1 = Left(pS, p - 1)
    oS2 = mID(pS, p + Len(pBrkChr))
Else
    oS1 = Trim(Left(pS, p - 1))
    oS2 = Trim(mID(pS, p + Len(pBrkChr)))
End If
Exit Function
E: Brk_Str_Both = True: ss.C cSub, cMod, "pS,pBrkChr,pNoTrim", pS, pBrkChr, pNoTrim
End Function
Function Brk_Str_1ForS2(oS1$, oS2$, pS$, pBrkChr$, Optional pNoTrim As Boolean = False) As Boolean
'Aim: Brk {pS} (format <oS1>[<pBrkChr><oS2>]).  If oS1 is missing, Set oS1="" & oS2=pS
Const cSub$ = "Brk_Str_1Or2"
If pBrkChr = "" Then ss.A 1, "BrkChr must be given": GoTo E
Dim p%: p = InStr(pS, pBrkChr)
If p = 0 Then
    If pNoTrim Then
        oS2 = pS
    Else
        oS2 = Trim(pS)
    End If
    oS1 = ""
    Exit Function
End If

If pNoTrim Then
    oS1 = Left(pS, p - 1)
    oS2 = mID(pS, p + Len(pBrkChr))
Else
    oS1 = Trim(Left(pS, p - 1))
    oS2 = Trim(mID(pS, p + Len(pBrkChr)))
End If
Exit Function
E: Brk_Str_1ForS2 = True: ss.B cSub, cMod, "pS,pBrkChr,pNoTrim", pS, pBrkChr, pNoTrim
End Function
Function Brk_Str_0Or2(oS1, oS2, pS$, pBrkChr$, Optional pNoTrim As Boolean = False) As Boolean
'Aim: Brk {pS} (format <oS1>[<pBrkChr><oS2>]).
Const cSub$ = "Brk_Str_1Or2"
On Error GoTo R
If pBrkChr = "" Then ss.A 1, "BrkChr must be given": GoTo E
Dim p%: p = InStr(pS, pBrkChr)
If p = 0 Then
    If pNoTrim Then
        oS1 = pS
    Else
        oS1 = Trim(pS)
    End If
    oS2 = ""
    Exit Function
End If

If pNoTrim Then
    oS1 = Left(pS, p - 1)
    oS2 = mID(pS, p + Len(pBrkChr))
Else
    oS1 = Trim(Left(pS, p - 1))
    oS2 = Trim(mID(pS, p + Len(pBrkChr)))
End If
Exit Function
R: ss.R
E: Brk_Str_0Or2 = True: ss.B cSub, cMod, "pS,pBrkChr,pNoTrim", pS, pBrkChr, pNoTrim
End Function
Function Brk_Str_1Or2(oS1$, oS2$, pS$, pBrkChr$, Optional pNoTrim As Boolean = False) As Boolean
'Aim: Brk {pS} (format <oS1>[<pBrkChr><oS2>]).
Const cSub$ = "Brk_Str_1Or2"
If pBrkChr = "" Then ss.A 1, "BrkChr must be given": GoTo E
If Trim(pS) = "" Then ss.A 2, "pS must be given": GoTo E
Dim p%: p = InStr(pS, pBrkChr)
If p = 0 Then
    If pNoTrim Then
        oS1 = pS
    Else
        oS1 = Trim(pS)
    End If
    oS2 = ""
    Exit Function
End If

If pNoTrim Then
    oS1 = Left(pS, p - 1)
    oS2 = mID(pS, p + Len(pBrkChr))
Else
    oS1 = Trim(Left(pS, p - 1))
    oS2 = Trim(mID(pS, p + Len(pBrkChr)))
End If
Exit Function
E: Brk_Str_1Or2 = True: ss.B cSub, cMod, "pS,pBrkChr,pNoTrim", pS, pBrkChr, pNoTrim
End Function
Function Brk_Str_1For2(oS1$, oS2$, pS$, pBrkChr$, Optional pNoTrim As Boolean = False) As Boolean
'Aim: Brk {pS} (format <oS1>[<pBrkChr><oS2>]).  If <oS2> is missing, set oS2=oS1
Const cSub$ = "Brk_Str_1For2"
If pBrkChr = "" Then ss.A 1, "BrkChr must be given": GoTo E
Dim p%: p = InStr(pS, pBrkChr)
If p = 0 Then
    If pNoTrim Then
        oS1 = pS
        oS2 = pS
    Else
        oS1 = Trim(pS)
        oS2 = oS1
    End If
    Exit Function
End If

If pNoTrim Then
    oS1 = Left(pS, p - 1)
    oS2 = mID(pS, p + Len(pBrkChr))
Else
    oS1 = Trim(Left(pS, p - 1))
    oS2 = Trim(mID(pS, p + Len(pBrkChr)))
End If
Exit Function
E: Brk_Str_1For2 = True: ss.B cSub, cMod, "pS,pBrkChr,pNoTrim", pS, pBrkChr, pNoTrim
End Function
Function Brk_Str_To3Seg(oS1, oS2, oS3, pS$, Optional pBrkChr$ = ":", Optional pNoTrim As Boolean = False) As Boolean
Const cSub$ = "Brk_Str_To3Seg"
Dim A$
If jj.Brk_Str_0Or2(oS1, A, pS, pBrkChr, pNoTrim) Then ss.A 1: GoTo E
Brk_Str_To3Seg = jj.Brk_Str_0Or2(oS2, oS3, A, pBrkChr, pNoTrim)
Exit Function
R: ss.R
E: Brk_Str_To3Seg = True: ss.C cSub, cMod, "pS,pBrkChr,pNoTrim", pS, pBrkChr, pNoTrim
End Function
Function Brk_Str_To4Seg(oS1, oS2, oS3, oS4, pS$, Optional pBrkChr$ = ":", Optional pNoTrim As Boolean = False) As Boolean
Const cSub$ = "Brk_Str2Seg4"
Dim A$
If jj.Brk_Str_To3Seg(oS1, oS2, A, pS, pBrkChr, pNoTrim) Then ss.A 1: GoTo E
Brk_Str_To4Seg = jj.Brk_Str_0Or2(oS3, oS4, A, pBrkChr, pNoTrim)
Exit Function
R: ss.R
E: Brk_Str_To4Seg = True: ss.B cSub, cMod, "pS,pBrkChr,pNoTrim", pS, pBrkChr, pNoTrim
End Function
Function Brk_Str_To5Seg(oS1, oS2, oS3, oS4, oS5, pS$, Optional pBrkChr$ = ":", Optional pNoTrim As Boolean = False) As Boolean
Const cSub$ = "Brk_Str_To5Seg"
Dim A$
If jj.Brk_Str_To4Seg(oS1, oS2, oS3, A, pS, pBrkChr, pNoTrim) Then ss.A 1: GoTo E
Brk_Str_To5Seg = jj.Brk_Str_0Or2(oS4, oS5, A, pBrkChr, pNoTrim)
Exit Function
R: ss.R
E: Brk_Str_To5Seg = True: ss.B cSub, cMod, "pS,pBrkChr,pNoTrim", pS, pBrkChr, pNoTrim
End Function
Function Brk_Str2Map(oMap As tMap, pS$, Optional pBrkChr$ = "=", Optional pNoTrim As Boolean = False) As Boolean
With oMap
    Brk_Str2Map = jj.Brk_Str_1For2(.F1, .F2, pS, pBrkChr$, pNoTrim)
End With
End Function
