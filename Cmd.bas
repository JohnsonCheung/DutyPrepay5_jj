Attribute VB_Name = "Cmd"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".Cmd"
Function CmdClick() As Boolean
Const cSub$ = "CmdClick"
'Find mActiveCmdBarCtl
On Error GoTo R
Dim mActiveCmdBarCtl As CommandBarControl:  Set mActiveCmdBarCtl = Application.CommandBars.ActionControl
If jj.IsNothing(mActiveCmdBarCtl) Then Exit Function
If mActiveCmdBarCtl.Type <> msoControlButton Then Exit Function
Dim mNmToolBar$: mNmToolBar = mActiveCmdBarCtl.Parent.Name
Dim mCmd$:       mCmd = mActiveCmdBarCtl.Parameter

'React on cmdBack
Select Case mCmd
Case "cmdBack"
    If Application.CurrentObjectType = acQuery Then DoCmd.Close: Exit Function
    If Application.CurrentObjectType = acTable Then DoCmd.Close: Exit Function
    If Forms.Count = 1 Then Fct.Quit: Exit Function
    On Error Resume Next
    DoCmd.Close
    If Forms.Count = 1 Then Forms(1).SetFocus
    Exit Function
Case "SetVer"
    If jj.Set_Ver Then ss.A 1: GoTo E
    Exit Function
End Select
CmdClick = jj.Run_Prc(mNmToolBar & "_" & mCmd)
Exit Function
R: ss.R

E:
: ss.B cSub, cMod
    CmdClick = True
End Function
Function CmdOpn_Calendar() As Boolean
Const cSub$ = "CmdOpn_Calendar"
Dim mActiveCmdBarCtl As CommandBarControl: Set mActiveCmdBarCtl = Application.CommandBars.ActionControl
If mActiveCmdBarCtl.Type <> msoControlButton Then Exit Function
Dim mParam$: mParam = mActiveCmdBarCtl.Parameter
'
Dim mFfn$: mFfn = jj.Sdir_Tp & "SPL Company Calendar FY" & mParam & ".xls"
If VBA.Dir(mFfn) = "" Then ss.A 1, "Calendar file not found": GoTo E
'
Dim mXls As New Excel.Application
mXls.Visible = True
mXls.Workbooks.Open mFfn
Exit Function
E:
: ss.B cSub, cMod
    CmdOpn_Calendar = True
End Function
Function zCommon_FY06() As Boolean: zCommon_FY06 = jj.Opn_Calendar(6): End Function
Function zCommon_FY07() As Boolean: zCommon_FY07 = jj.Opn_Calendar(7): End Function
Function zCommon_FY08() As Boolean: zCommon_FY08 = jj.Opn_Calendar(8): End Function
Function zCommon_FY09() As Boolean: zCommon_FY09 = jj.Opn_Calendar(9): End Function
Function zCommon_FY10() As Boolean: zCommon_FY10 = jj.Opn_Calendar(10): End Function
Function zCommon_FY11() As Boolean: zCommon_FY11 = jj.Opn_Calendar(11): End Function
Function zCommon_FY12() As Boolean: zCommon_FY12 = jj.Opn_Calendar(12): End Function
Function zCommon_FY13() As Boolean: zCommon_FY13 = jj.Opn_Calendar(13): End Function
Function zCommon_FY14() As Boolean: zCommon_FY14 = jj.Opn_Calendar(14): End Function
Function zCommon_FY15() As Boolean: zCommon_FY15 = jj.Opn_Calendar(15): End Function
