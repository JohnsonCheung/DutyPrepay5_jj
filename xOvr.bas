Attribute VB_Name = "xOvr"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xOvr"
Function Ovr_Wrt(pFfn$, pOvrWrt As Boolean) As Boolean
Const cSub$ = "Ovr_Wrt"
If VBA.Dir(pFfn) = "" Then Exit Function
If Not pOvrWrt Then ss.A 1, "File Exist": GoTo E
If jj.Dlt_Fil(pFfn) Then ss.A 2: GoTo E
Exit Function
R: ss.R
E: Ovr_Wrt = True: ss.B cSub, cMod, "pFfn,pOvrWrt", pFfn, pOvrWrt
End Function
