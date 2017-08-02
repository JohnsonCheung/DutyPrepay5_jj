Attribute VB_Name = "xSel"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xSel"
Function Sel_Brand() As Boolean
Sel_Brand = jj.Opn_Frm("frmSelBrand")
End Function
Function Sel_BrandEnv() As Boolean
Sel_BrandEnv = jj.Opn_Frm("frmSelBrandEnv")
End Function
Function Sel_FY() As Boolean
Sel_FY = jj.Opn_Frm("frmSelFY")
End Function

