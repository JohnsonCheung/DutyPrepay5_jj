VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "d_Pgm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".d_Pgm"
Public x_NmPrj$, x_Nmm$, x_NmPrc$, x_NmTypRet$, x_TypFct As eTypFct, x_IsPrivate As Boolean, x_Aim$, x_PrcBody
Function CrtTbl(Optional pNmt$ = "#OldPgm") As Boolean
Const cSub$ = "CrtTbl"
If jj.Crt_Tbl_FmLoFld(pNmt, "NmPrj Text 50, Nmm Text 50, NmPrc Text 50, NmTypRet Text 50,TypFct Byte, IsPrivate YesNo, Aim Memo, PrcBody Memo") Then ss.A 1: GoTo E
Exit Function
R: ss.R

E:
: ss.B cSub, cMod, "pNmt", pNmt
    CrtTbl = True
End Function
Function Ins(pRs As DAO.Recordset) As Boolean
Const cSub$ = "Ins"
With pRs
    Dim mLenNmTypRet%: mLenNmTypRet = .Fields("NmTypRet").Size
    .AddNew
    !NmPrj = x_NmPrj
    !Nmm = x_Nmm
    !NmPrc = x_NmPrc
    !NmTypRet = Left(x_NmTypRet, mLenNmTypRet)
    !TypFct = x_TypFct
    !IsPrivate = x_IsPrivate
    !Aim = x_Aim
    !PrcBody = x_PrcBody
    .Update
End With
End Function

