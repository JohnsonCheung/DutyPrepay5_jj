VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "d_QsT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit
Const cMod$ = cLib & ".d_QsT"
Public x_Fb$, x_NmQs$, x_NmTbl$, x_LnFld$
Function CrtTbl(Optional pNmt$ = "#OldQsT") As Boolean
Const cSub$ = "CrtTbl"
If jj.Crt_Tbl_FmLoFld(pNmt, "Fb Text 255, NmQs Text 255, NmTbl Text 255, LnFld Memo") Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: CrtTbl = True
: ss.B cSub, cMod, "pNmt", pNmt
End Function
Function InsAy(pAyOldQsT() As d_QsT, pNmtOldQsT$) As Boolean
'Aim: Insert pAyOldQsT to pNmtOldQsT$
Const cSub$ = "InsAy"
Dim J%, mN%: mN = jj.Siz_AyOldQsT(pAyOldQsT)
If mN = 0 Then Exit Function
On Error GoTo R
Dim mRs As DAO.Recordset: Set mRs = CurrentDb.TableDefs(pNmtOldQsT).OpenRecordset
For J = 0 To mN - 1
    If pAyOldQsT(J).Ins(mRs) Then ss.A 1: GoTo E
Next
GoTo X
R: ss.R
E: InsAy = True
: ss.B cSub, cMod, "pAyOldQsT(),pNmtOldQsT", "..", pNmtOldQsT
X:
    jj.Cls_Rs mRs
End Function
Function Ins(pRs As DAO.Recordset) As Boolean
Const cSub$ = "Ins"
On Error GoTo R
With pRs
    .AddNew
    !Fb = x_Fb
    !NmQs = x_NmQs
    !NmTbl = x_NmTbl
    !LnFld = x_LnFld
    .Update
End With
Exit Function
R: ss.R
E: Ins = True
: ss.B cSub, cMod
End Function
Function WrtHdr(pF As Byte) As Boolean
Const cSub$ = "WrtHdr"
On Error GoTo R
Write #pF, "Fb", "NmQs", "NmTbl", "LnFld"
Exit Function
R: ss.R
E: WrtHdr = True
: ss.B cSub, cMod, "pF", pF
End Function
Function Wrt(pF As Byte) As Boolean
Const cSub$ = "Wrt"
On Error GoTo R
Write #pF, x_Fb, x_NmQs, x_NmTbl, x_LnFld
Exit Function
R: ss.R
E: Wrt = True
: ss.B cSub, cMod, "pF", pF
End Function

