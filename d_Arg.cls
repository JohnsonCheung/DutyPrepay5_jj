VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "d_Arg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".d_Arg"
Public x_NmArg$, x_NmTypArg$, x_IsAy As Boolean, x_IsPrmAy As Boolean, x_IsOpt As Boolean, x_IsByVal As Boolean, x_DftVal
Function CrtTbl(Optional pNmt$ = "#OldArg") As Boolean
Const cSub$ = "CrtTbl"
If jj.Crt_Tbl_FmLoFld(pNmt, "NmPrj Text 50,Nmm Text 50,NmPrc Text 50, SnoArg Byte, NmArg Text 50, NmTypArg Text 50,IsAy YesNo, IsPrmAy YesNo, IsOpt YesNo, IsByVal YesNo, DftVal Text 50") Then ss.A 1: GoTo E
Exit Function
R: ss.R

E:
: ss.B cSub, cMod, "pNmt", pNmt
    CrtTbl = True
End Function
Function InsAy(pRs As DAO.Recordset, pNmPrj$, pNmm$, pNmPrc$, pAyDArg() As d_Arg) As Boolean
'Aim: Insert pAyDArg to pRs
Const cSub$ = "InsAy"
Dim J%, mN%: mN = jj.Siz_AyDArg(pAyDArg)
If mN = 0 Then Exit Function
On Error GoTo R
For J = 0 To mN - 1
    If pAyDArg(J).Ins(pRs, pNmPrj, pNmm, pNmPrc, J) Then ss.A 1: GoTo E
Next
Exit Function
R: ss.R

E:
: ss.B cSub, cMod, "pRs,pNmPrj,pNmm,pNmPrc,pAyDArg()", jj.ToStr_Rs_NmFld(pRs), pNmPrj, pNmm, pNmPrc, jj.ToStr_AyDArg(pAyDArg)
    InsAy = True
End Function
Function Ins(pRs As DAO.Recordset, pNmPrj$, pNmm$, pNmPrc$, pSnoArg%) As Boolean
Const cSub$ = "Ins"
With pRs
    Dim mLenNmArg%: mLenNmArg = .Fields("NmArg").Size
    Dim mLenNmTypArg%: mLenNmTypArg = .Fields("NmTypArg").Size
    Dim mLenDftVal%: mLenDftVal = .Fields("DftVal").Size
    .AddNew
    !NmPrj = pNmPrj
    !Nmm = pNmm
    !NmPrc = pNmPrc
    !NmArg = Left(x_NmArg, mLenNmArg)
    !SnoArg = pSnoArg
    !NmTypArg = Left(x_NmTypArg, mLenNmTypArg)
    !IsAy = x_IsAy
    !IsPrmAy = x_IsPrmAy
    !IsOpt = x_IsOpt
    !IsByVal = x_IsByVal
    !DftVal = Left(x_DftVal, mLenDftVal)
    .Update
End With
End Function
