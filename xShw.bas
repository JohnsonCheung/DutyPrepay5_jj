Attribute VB_Name = "xShw"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xShw"
Function Shw_QryChk(pAnqChk$()) As Boolean
'Aim: Assume all {pAnqChk} ends with _ChkXXX.
'     - Set the Sql of the each of the {pAnqChk} with Select * from [#ChkXXX]
'     - Check Table #ChkXXX exist
'           Delete them all if they have no record.
'           If record exist, use {pAnqChk} to open the table [#ChkXXX] and return error
Const cSub$ = "Shw_QryChk"
On Error GoTo R
Dim N%: N = jj.Siz_Ay(pAnqChk): If N = 0 Then Exit Function
Dim J%, mIsEr As Boolean: mIsEr = False
For J = 0 To N - 1
    Dim p%: p = InStrRev(pAnqChk(J), "_#Chk"): If p = 0 Then ss.A 1, "Given pAnqChk must contain _#ChkXXX from end", , "J", J: GoTo E
    Dim mNmtChk$: mNmtChk = mID(pAnqChk(J), p + 1)
    If jj.IsTbl(mNmtChk) Then
        Dim mRecCnt&: If jj.Fnd_RecCnt_ByNmtq(mRecCnt, mNmtChk) Then ss.A 3: GoTo E
        If mRecCnt = 0 Then
            If jj.Dlt_Tbl(mNmtChk) Then ss.A 5: GoTo E
        Else
            mIsEr = True
            DoCmd.OpenQuery pAnqChk(J)
        End If
    End If
Next
If mIsEr Then ss.A 4, "#Chk* table exist means error": GoTo E
Exit Function
R: ss.R
E: Shw_QryChk = True: ss.B cSub, cMod, "pAnqChk$", jj.ToStr_Ays(pAnqChk)
End Function
Function Shw_AllDta(pWs As Worksheet) As Boolean
jj.Shw_AllCols pWs
On Error Resume Next
pWs.ShowAllData
End Function
Function Shw_AllCols(pWs As Worksheet) As Boolean
pWs.Application.ScreenUpdating = False
Dim mOutLine As OutLine: Set mOutLine = pWs.OutLine
On Error Resume Next
Dim J As Byte
For J = 0 To 8
    mOutLine.ShowLevels , J
Next
pWs.Application.ScreenUpdating = True
End Function
Function Shw_Dbg(pSub$, pMod$ _
            , Optional pLp$ = "" _
            , Optional pV0 _
            , Optional pV1 _
            , Optional pV2 _
            , Optional pV3 _
            , Optional pV4 _
            , Optional pV5 _
            , Optional pV6 _
            , Optional pV7 _
            , Optional pV8 _
            , Optional pV9 _
            , Optional pV10 _
            , Optional pV11 _
            , Optional pV12 _
            , Optional pV13 _
            , Optional pV14 _
            , Optional pV15 _
    ) As Boolean
'PrcDcl: Shw the PrcDcl of the {pSub} (Asumme pSub ends with _Tst) and debug variable
Const cSub$ = "Shw_Dbg"
On Error GoTo R
If Right(pSub, 4) <> "_Tst" Then ss.A 1, "Last 4 char of pSub must be _Tst": GoTo E
Dim p As Byte: p = InStr(pMod, "."): If p = 0 Then ss.A 2, "pMod must have .": GoTo E
Dim mNmPrc$: mNmPrc = Left(pSub, Len(pSub) - 4)
'
Dim mNmPrc_Full$: mNmPrc_Full = pMod & "." & mNmPrc
Debug.Print String(Len(mNmPrc_Full), "-")
Debug.Print mNmPrc_Full
Debug.Print String(Len(mNmPrc_Full), "-")
'
Dim mPrcDcl$, mMaxLin%
If Not jj.Fnd_PrcDcl(mPrcDcl, pMod, mNmPrc) Then
    If mPrcDcl <> "" Then
        mMaxLin = jj.Fnd_MaxLin(mPrcDcl)
        Debug.Print mPrcDcl
        Debug.Print String(mMaxLin, "-")
    End If
End If
If pLp <> "" Then
    Debug.Print jj.ToStr_LpAp(vbLf, pLp, pV0, pV1, pV2, pV3, pV4, pV5, pV6, pV7, pV8, pV9, pV10, pV11, pV12, pV13, pV14, pV15)
    Debug.Print
End If
jj.Shw_DbgWin
Exit Function
R: ss.R
E: Shw_Dbg = True: ss.C cSub, cMod, "pSub,pMod,pLp", pSub, pMod, pLp
End Function
Function Shw_Dbg_Tst() As Boolean
Const cSub$ = "Shw_Dbg_Tst"
jj.Shw_Dbg cSub, cMod, "VarA,VarB", 1, 2
End Function
Function Shw_Msg_ByAm(pPrcDcl$, pMsgNo As Byte, pSub$, pMod$, pTypMsg As eTypMsg, pTit$, pAm() As tMap) As Boolean
If g.gIsBch Then Exit Function
If g.gSilent Then Exit Function
Dim xTit$
Dim xMsg$
Dim xMsgBoxSty As VbMsgBoxStyle
If Not jj.SysCfg_IsDbg Then If pTypMsg = eTypMsg.eSeePrvMsg Then Exit Function

xMsgBoxSty = Fnd_MsgBoxSty(pTypMsg)
Dim mTit$
If pTit = "" Then
    mTit = jj.ToStr_TypMsg(pTypMsg)
    xTit = pMod & "." & pSub & "(" & pMsgNo & ")"
Else
    mTit = Replace(pTit, "|", vbLf)
    xTit = pMod & "." & pSub & "(" & pMsgNo & ") " & jj.ToStr_TypMsg(pTypMsg)
End If
If pPrcDcl <> "" Then pPrcDcl = vbLf & vbLf & pPrcDcl
xMsg = mTit & pPrcDcl & vbLf & vbLf & jj.ToStr_Am(pAm, , , "[]")
If MsgBox(xMsg, xMsgBoxSty, xTit) = vbYes Then GoTo E
Exit Function
E: Shw_Msg_ByAm = True
End Function
Function Shw_DbgWin() As Boolean
DoCmd.RunCommand acCmdDebugWindow
DoCmd.Maximize
End Function
Function Shw_Log() As Boolean
Const cSub$ = "Shw_Log"
Dim mFfnDbLog$: mFfnDbLog = jj.Sffn_DbLog
If jj.Crt_Qry("qryShwLog", jj.Fmt_Str("Select * from tblLog in '{0}' order by MsgId Desc", mFfnDbLog)) Then ss.A 1: GoTo E
If jj.Opn_Qry("qryShwLog") Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: Shw_Log = True: ss.B cSub, cMod
End Function
Function Shw_Lvl1_ByWb(pWb As Workbook) As Boolean
Dim iWs As Worksheet
For Each iWs In pWb.Sheets
    iWs.OutLine.ShowLevels 1, 1
Next
End Function
Function Shw_Sts(pS$) As Boolean
Debug.Print pS
SysCmd acSysCmdSetStatus, pS
End Function
Function Shw_ToolBar(pNmToolBar$, Optional pVisible As Boolean = True) As Boolean
Const cSub$ = "Shw_ToolBar"
On Error GoTo R
Dim mCmdBar As CommandBar: Set mCmdBar = Application.CommandBars(pNmToolBar)
mCmdBar.Visible = pVisible
Exit Function
R: ss.R
E: Shw_ToolBar = True: ss.B cSub, cMod, "pNmToolBar", pNmToolBar
End Function
Function Shw_ToolBarBtn(pNmToolBar$, pLoBtnPrm$, Optional pEnable As Boolean = True) As Boolean
Const cSub$ = "Shw_ToolBarBtn"
Dim iCtl As CommandBarControl, mFound As Boolean
For Each iCtl In Application.CommandBars(pNmToolBar).Controls
    If InStr(pLoBtnPrm, iCtl.Parameter) > 0 Then iCtl.Enabled = pEnable: mFound = True
Next
If Not mFound Then ss.A 1, "No given BtnPrm in ToolBar": ss.A 1: GoTo E
Exit Function
E: Shw_ToolBarBtn = True: ss.B cSub, cMod, "pNmToolBar,pLoBtnPrm", pNmToolBar, pLoBtnPrm
End Function
