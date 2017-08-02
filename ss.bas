Attribute VB_Name = "ss"
#Const Tst = True
Option Compare Text
Option Explicit
Const cMod$ = cLib & ".ss"
Public xMonMsg%
Public xMonMsgMatched As Boolean
Private xDbLog As Database
Public xMsgNo As Byte
Private xTit$
Private xTypMsg As eTypMsg
Private xLp$, xV0, xV1, xV2, xV3, xV4, xV5
Function ClsDbLog() As Boolean
jj.Cls_Db xDbLog
End Function
Function R() As Boolean
xMsgNo = 255
xTit = Err.Description
xMonMsgMatched = xMonMsg = Err.Number
xTypMsg = eException
End Function
Sub A(pMsgNo As Byte, Optional pTit$ = "", Optional pTypMsg As eTypMsg = eTypMsg.eSeePrvMsg _
            , Optional pLp$ = "" _
            , Optional pV0 _
            , Optional pV1 _
            , Optional pV2 _
            , Optional pV3 _
            , Optional pV4 _
            , Optional pV5 _
    )
xMsgNo = pMsgNo
xTit = pTit
xTypMsg = pTypMsg
xLp = pLp
xV0 = pV0
xV1 = pV1
xV2 = pV2
xV3 = pV3
xV4 = pV4
xV5 = pV5
'If pMsgNo <> 255 Then Stop
End Sub
Sub B(pSub$, pMod$ _
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
    )
Const cSub$ = "B"
Dim mAm1() As tMap: If jj.Cv_Lp16vToAm(mAm1, pLp, pV0, pV1, pV2, pV3, pV4, pV5, pV6, pV7, pV8, pV9, pV10, pV11, pV12, pV13, pV14, pV15) Then ss.A 1: GoTo E
Dim mAm2() As tMap: If jj.Cv_Lp16vToAm(mAm2, xLp, xV0, xV1, xV2, xV3, xV4, xV5) Then ss.A 2: GoTo E
Dim mAm() As tMap: If jj.Add_Am(mAm, mAm1, mAm2) Then ss.A 3: GoTo E
Dim mMsgId&: If zzLogMsg(mMsgId, xMsgNo, pSub, pMod, xTypMsg, xTit, mAm) Then ss.A 4: GoTo E
If ss.xMonMsgMatched Then Exit Sub
If xTypMsg = eTrc Then Exit Sub
If jj.g.gSilent Then Exit Sub
If jj.g.gIsBch Then Exit Sub
If xTypMsg = eUsrInfo Then zzShwMsg mMsgId: Exit Sub

'Dim mPrcDcl$: If jj.Fnd_PrcDcl(mPrcDcl, pMod, pSub) Then MsgBox "Cannot find PrcDlc for pMod=[" & pMod & "] pSub=[" & pSub & "]": GoTo E
Dim mPrcDcl$: jj.Fnd_PrcDcl mPrcDcl, pMod, pSub
If zzShwMsg(mMsgId, mPrcDcl) Then Stop
Exit Sub
E: ss.C cSub, cMod, cSub, "pSub,pMod,pLp,pV0,..", pSub, pMod, pLp, pV0, ".."
End Sub
Sub C(pSub$, pMod$ _
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
    )
Const cSub$ = "C"
If xMsgNo = 255 And xMonMsgMatched Then Exit Sub
Dim mAm1() As tMap: If jj.Cv_Lp16vToAm(mAm1, pLp, pV0, pV1, pV2, pV3, pV4, pV5, pV6, pV7, pV8, pV9, pV10, pV11, pV12, pV13, pV14, pV15) Then ss.A 1: GoTo E
Dim mAm2() As tMap: If jj.Cv_Lp16vToAm(mAm2, xLp, xV0, xV1, xV2, xV3, xV4, xV5) Then ss.A 2: GoTo E
Dim mAm() As tMap: If jj.Add_Am(mAm, mAm1, mAm2) Then ss.A 3: GoTo E
If ss.xMonMsgMatched Then Exit Sub
If xTypMsg = eTrc Then Exit Sub
If jj.g.gSilent Then Exit Sub
If jj.g.gIsBch Then Exit Sub
Dim mPrcDcl$: If jj.Fnd_PrcDcl(mPrcDcl, pMod, pSub) Then ss.A 5: GoTo E
If jj.Shw_Msg_ByAm(mPrcDcl, xMsgNo, pSub, pMod, xTypMsg, xTit, mAm) Then Stop
Exit Sub
E: MsgBox "Critical error in ss.C": Stop
End Sub
Sub xx(pMsgNo As Byte, pSub$, pMod$ _
            , Optional pTypMsg As eTypMsg = eTypMsg.eSeePrvMsg _
            , Optional pTit$ = "" _
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
    )
Const cSub$ = "xx"
If jj.g.gSilent Then Exit Sub
Dim mAm() As tMap: If jj.Cv_Lp16vToAm(mAm, pLp, pV0, pV1, pV2, pV3, pV4, pV5, pV6, pV7, pV8, pV9, pV10, pV11, pV12, pV13, pV14, pV15) Then ss.A 1: GoTo E
Dim mMsgId&: If zzLogMsg(mMsgId, pMsgNo, pSub, pMod, pTypMsg, pTit, mAm) Then ss.A 2: GoTo E
If pTypMsg = eTrc Then Exit Sub
If jj.g.gSilent Then Exit Sub
Dim mPrcDcl$: If jj.Fnd_PrcDcl(mPrcDcl, pMod, pSub) Then ss.A 5: GoTo E
If zzShwMsg(mMsgId, mPrcDcl) Then Stop
Exit Sub
E: ss.C cSub, cMod, cSub, "pSub,pMod,pTypMsg,pTit,pLp,pV0,..", pSub, pMod, jj.ToStr_TypMsg(pTypMsg), pTit, pLp, pV0, ".."
End Sub
#If Tst Then
Function A_Tst() As Boolean
Const cSub$ = "A_Tst"
ss.A 1, "Test", eRunTimErr, "abc,date", 1, Now
E: A_Tst = True: ss.B cSub, cMod, "xx,yy", 3, 4
End Function
#End If
Private Function zzCrtLogDb() As Boolean
Dim mFfnFm$: mFfnFm = jj.Sdir_PgmObj & "Template_Log.Mdb"
Dim mFfnTo$: mFfnTo = jj.Sffn_DbLog
On Error GoTo R
VBA.FileCopy mFfnFm, mFfnTo
Exit Function
R: ss.R
    MsgBox "Cannot copy file" & vbLf & "FfnFm=[" & mFfnFm & "]" & vbLf & "FfnTo=[" & mFfnTo & "]" & vbLf & "Err=[" & Err.Description & "]", vbCritical, "Critical"
    zzCrtLogDb = True
End Function
Private Function zzDbLog() As Database
If jj.IsNothing(xDbLog) Then If ss.zzOpnLogDb() Then Stop
Set zzDbLog = xDbLog
End Function
Private Function zzLogMsg(oMsgId&, pMsgNo As Byte, pSub$, pMod$, pTypMsg As eTypMsg, pTit$, pAm() As tMap) As Boolean
With zzDbLog.TableDefs("tblLog").OpenRecordset
    .AddNew
    oMsgId = !MsgId
    !DteTim = Now
    !Mod = pMod
    !Sub = pSub
    !MsgNo = pMsgNo
    !TypMsg = pTypMsg
    !Tit = Replace(pTit, "|", vbCrLf)
    !NmLst = jj.ToStr_AmF1(pAm)
    Dim N%: N = jj.Siz_Am(pAm)
    Do
        If N > 0 Then !V0 = pAm(0).F1 & "=" & pAm(0).F2 Else Exit Do
        If N > 1 Then !V1 = pAm(1).F1 & "=" & pAm(1).F2 Else Exit Do
        If N > 2 Then !V2 = pAm(2).F1 & "=" & pAm(2).F2 Else Exit Do
        If N > 3 Then !V3 = pAm(3).F1 & "=" & pAm(3).F2 Else Exit Do
        If N > 4 Then !V4 = pAm(4).F1 & "=" & pAm(4).F2 Else Exit Do
        If N > 5 Then !V5 = pAm(5).F1 & "=" & pAm(5).F2 Else Exit Do
        If N > 6 Then !V6 = pAm(6).F1 & "=" & pAm(6).F2 Else Exit Do
        If N > 7 Then !V7 = pAm(7).F1 & "=" & pAm(7).F2 Else Exit Do
        If N > 8 Then !V8 = pAm(8).F1 & "=" & pAm(8).F2 Else Exit Do
        If N > 9 Then !V9 = pAm(9).F1 & "=" & pAm(9).F2 Else Exit Do
        If N > 10 Then !V10 = pAm(10).F1 & "=" & pAm(10).F2 Else Exit Do
        If N > 11 Then !V11 = pAm(11).F1 & "=" & pAm(11).F2 Else Exit Do
        If N > 12 Then !V12 = pAm(12).F1 & "=" & pAm(12).F2 Else Exit Do
        If N > 13 Then !V13 = pAm(13).F1 & "=" & pAm(13).F2 Else Exit Do
        If N > 15 Then !V14 = pAm(14).F1 & "=" & pAm(14).F2 Else Exit Do
        If N > 15 Then !V15 = pAm(15).F1 & "=" & pAm(15).F2
        Exit Do
    Loop Until False
    .Update
    .Close
End With
End Function
Private Function zzOpnLogDb() As Boolean
Const cSub$ = "zzOpnLogDb"
On Error GoTo R
Dim mFfnDbLog$: mFfnDbLog = jj.Sffn_DbLog
If VBA.Dir(mFfnDbLog) = "" Then If zzCrtLogDb() Then ss.A 1: GoTo E
Set xDbLog = g.gDbEng.OpenDatabase(mFfnDbLog)
If Not jj.IsTbl("tblLog", xDbLog) Then ss.A 2: GoTo E
Exit Function
R: ss.R
E: zzOpnLogDb = True: ss.C cSub, cMod, "mFfnDbLog", mFfnDbLog
End Function
Private Function zzShwMsg(pMsgId&, Optional pPrcDcl$ = "") As Boolean
If g.gIsBch Then Exit Function
If g.gSilent Then Exit Function
Dim xTit$
Dim xMsg$
Dim xMsgBoxSty As VbMsgBoxStyle
With zzDbLog.OpenRecordset("Select * from tblLog where MsgId=" & pMsgId)
    If .EOF Then MsgBox ("MsgId=" & pMsgId & " not found in tblLog"): .Close: GoTo E
    If Not jj.SysCfg_IsDbg Then If !TypMsg = eTypMsg.eSeePrvMsg Then .Close: Exit Function

    xMsgBoxSty = Fnd_MsgBoxSty(!TypMsg)
    Dim mTit$
    If !Tit = "" Then
        mTit = jj.ToStr_TypMsg(!TypMsg)
        xTit = !Mod & "." & !Sub & "(" & !MsgNo & ")"
    Else
        mTit = Replace(!Tit, "|", vbLf)
        xTit = !Mod & "." & !Sub & "(" & !MsgNo & ") " & jj.ToStr_TypMsg(!TypMsg)
    End If
    If pPrcDcl <> "" Then pPrcDcl = vbLf & vbLf & pPrcDcl
    xMsg = mTit & pPrcDcl & vbLf & vbLf & jj.ToStr_LpAp(vbLf, Nz(!NmLst.Value, ""), !V0.Value, !V1.Value, !V2.Value, !V3.Value, !V4.Value, !V5.Value, !V6.Value, !V7.Value, !V8.Value, !V9.Value, !V10.Value, !V11.Value, !V12.Value, !V13.Value, !V14.Value, !V15.Value)
    .Close
End With
If MsgBox(xMsg, xMsgBoxSty, xTit) = vbYes Then GoTo E
Exit Function
E: zzShwMsg = True
End Function



