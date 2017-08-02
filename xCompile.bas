Attribute VB_Name = "xCompile"
Option Compare Text
Option Explicit
Const cMod$ = cLib & ".xCompile"
Function Compile_Qry(pNmq$) As Boolean
'Aim: Compile a query
DoCmd.SetWarnings False
DoCmd.OpenQuery pNmq, acViewDesign
DoCmd.RunCommand acCmdSQLView
DoCmd.RunCommand acCmdCopy
DoCmd.RunCommand acCmdPaste
DoCmd.RunCommand acCmdClose
DoCmd.SetWarnings True
End Function
Function Compile_Xls(pXls As Excel.Application) As Boolean
Const cSub$ = "Compile_Acs"
On Error GoTo R
'
Stop
Exit Function
R: ss.R
E: Compile_Xls = True: ss.B cSub, cMod, "pXls", jj.ToStr_Xls(pXls)
End Function
Function Compile_Acs(pAcs As Access.Application) As Boolean
Const cSub$ = "Compile_Acs"
On Error GoTo R
pAcs.RunCommand acCmdCompileAndSaveAllModules
Exit Function
R: ss.R
E: Compile_Acs = True: ss.B cSub, cMod, "pAcs", jj.ToStr_Acs(pAcs)
End Function
Function Compile_App(pApp As Object) As Boolean
Const cSub$ = "Compile_App"
On Error GoTo R
Select Case TypeName(pApp)
Case "Microsoft Access":    Compile_App = jj.Compile_Xls(pApp)
Case "Microsoft Excel":     Compile_App = jj.Compile_Xls(pApp)
Case Else: ss.A 1, "Unexpected application", , "Expected Type", "Excel,Access": GoTo E
End Select
Exit Function
R: ss.R
E: Compile_App = True: ss.B cSub, cMod, "TypeName(pApp)", TypeName(pApp)
End Function
