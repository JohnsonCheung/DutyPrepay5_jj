Attribute VB_Name = "xGo"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xGo"
Function Go_Line(pNmm$, pTypObj As Access.AcObjectType, pL&) As Boolean
Dim mMod As Module
Select Case pTypObj
Case acForm
    Call DoCmd.OpenForm(pNmm, acDesign)
    Set mMod = Forms(pNmm).Module
Case acModule
    Call DoCmd.OpenModule(pNmm)
    Set mMod = Modules(pNmm)
Case Else
    Exit Function
End Select
Dim mCodePane As VBIDE.CodePane
For Each mCodePane In Application.VBE.CodePanes
    If UCase(mCodePane.CodeModule.Name) = UCase(pNmm) Then Exit For
Next
Call mCodePane.SetSelection(pL, 1, pL, 1)
mCodePane.Show
mCodePane.Window.SetFocus
End Function
Function Go_QryDef(pNmq$) As Boolean
Const cSub$ = "Go_QryDef"
On Error GoTo R
DoCmd.OpenQuery pNmq, acViewDesign
On Error GoTo 0
Exit Function
R: ss.R
E: Go_QryDef = True: ss.B cSub, cMod, "pNmq", pNmq
End Function
Function Go_QryRmk(pNmq$, pRmk$) As Boolean
jj.Set_Prp pNmq, acQuery, "Description", pRmk
End Function
Function Go_QryView(pNmq$) As Boolean
Const cSub$ = "Go_QryView"
On Error GoTo R
DoCmd.OpenQuery pNmq, , acReadOnly
On Error GoTo 0
Exit Function
R: ss.R
E: Go_QryView = True: ss.B cSub, cMod, "pNmq", pNmq
End Function
Function Go_Rec(Optional pWhere As AcRecord = acNext) As Boolean
On Error GoTo E
DoCmd.GoToRecord , , pWhere
Exit Function
E: Go_Rec = True
End Function
