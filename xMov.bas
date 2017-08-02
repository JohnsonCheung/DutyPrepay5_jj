Attribute VB_Name = "xMov"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xMov"
Function Mov_Fil(pDirFm$, pDirTo$, pFnFm$, Optional pFnTo$ = "") As Boolean
Const cSub$ = "Mov_Fil"
If pFnTo = "" Then pFnTo = pFnFm
If VBA.Dir(pDirFm & pFnFm) = "" Then ss.A 1, "From file not exist!", , "pDirTo,pFnFm,pFnTo", pDirTo, pFnFm, pFnTo: GoTo E
If VBA.Dir(pDirTo & pFnTo) <> "" Then ss.A 2, "To file exist", , "pDirFm,pFnFm,pFnTo", pDirFm, pFnFm, pFnTo: GoTo E
On Error GoTo R
gFso.MoveFile pDirFm & pFnFm, pDirTo & pFnTo
If VBA.Dir(pDirTo & pFnTo) = "" Then ss.A 1, "After move, the to file not exist": GoTo E
Exit Function
R: ss.R
E: Mov_Fil = True: ss.B cSub, cMod, "pDirFm,pFnFm,pFnTo", pDirFm, pFnFm, pFnTo
End Function
