Attribute VB_Name = "xUsrPrf"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xUsrPrf"
Private xUsrPrf As tUsrPrf
Private Const cLnItm$ = "Usr,Dpt,Fy,Env,Lvl,Brand"
'Debug.Print jj.Fmt_Str_Repeat_Lv("Nm{B} As String", cLnItm, vbLf)
'Debug.Print jj.Fmt_Str_Repeat_Lv("{B} As Long", cLnItm, vbLf)
Private Type tUsrPrf
    NmUsr As String
    NmDpt As String
    NmFy As String
    NmEnv As String
    NmLvl As String
    NmBrand As String
    Usr As Long
    Dpt As Long
    Fy As Long
    Env As Long
    Lvl As Long
    Brand As Long
End Type
'=====================================
'Debug.Print jj.Fmt_Str_Repeat_Lv("Function UsrPrf_Nm{B}$(): UsrPrf_Nm{B} = xx.Nm{B}: End Function", "Usr,Dpt,Fy,Env,Lvl,Brand", vbLf)
Function UsrPrf_NmUsr$(): UsrPrf_NmUsr = xx.NmUsr: End Function
Function UsrPrf_NmDpt$(): UsrPrf_NmDpt = xx.NmDpt: End Function
Function UsrPrf_NmFy$(): UsrPrf_NmFy = xx.NmFy: End Function
Function UsrPrf_NmEnv$(): UsrPrf_NmEnv = xx.NmEnv: End Function
Function UsrPrf_NmLvl$(): UsrPrf_NmLvl = xx.NmLvl: End Function
Function UsrPrf_NmBrand$(): UsrPrf_NmBrand = xx.NmBrand: End Function
'Debug.Print jj.Fmt_Str_Repeat_Lv("Function Nm{B}$(): Nm{B} = xx.{B}: End Function", "Usr,Dpt,Fy,Env,Lvl,Brand", vbLf)
Function UsrPrf_Usr&(): UsrPrf_Usr = xx.Usr: End Function
Function UsrPrf_Dpt&(): UsrPrf_Dpt = xx.Dpt: End Function
Function UsrPrf_Fy&(): UsrPrf_Fy = xx.Fy: End Function
Function UsrPrf_Env&(): UsrPrf_Env = xx.Env: End Function
Function UsrPrf_Lvl&(): UsrPrf_Lvl = xx.Lvl: End Function
Function UsrPrf_Brand&(): UsrPrf_Brand = xx.Brand: End Function
Function UsrPrf_Login(pNmUsr$) As Boolean
'Aim: Login by {pNmUsr}
Const cSub$ = "UsrPrf_Login"
Dim mPwd$: If jj.Fnd_ValFmSql(mPwd, "Select Password from tblUsr where NmUsr='" & pNmUsr & cQSng) Then ss.A 1: GoTo E
If jj.UsrPrf_ValidateLogin(pNmUsr, mPwd) Then ss.A 2: GoTo E
Exit Function
E: UsrPrf_Login = True: ss.B cSub, cMod, "pNmUsr", pNmUsr
End Function
#If Tst Then
Function UsrPrf_Login_Tst() As Boolean
If UsrPrf_Login("Johnson") Then Stop
End Function
#End If
Sub LetXXX()
'Property Let UsrPrf_{B}(p{B}&)
'Const cSub$ = "UsrPrf_{B}"
'If jj.Run_Sql_ByDbExec(jj.Fmt_Str("Update tblUsr SET {B}={B} Where Usr={1}", p{B}, xx.UsrPrf_Usr), CodeDb) Then ss.A 1: GoTo E
'Dim mA$: If jj.Fnd_ValFmSql(mA, jj.Fmt_Str("Select Nm{B} from tbl{B} where {B}={B}", p{B}), CodeDb) Then ss.A 2: GoTo E
'xUsrPrf.Nm{B} = mA
'xUsrPrf.{B} = p{B}
'Exit Function
'E: ss.B cSub, cMod, "p{B}", p{B}
'End Property
End Sub
Private Function Gen_LetXXX() As Boolean
Dim mA$: If jj.Fnd_ResStr(mA, "LetXXX", True) Then ss.A 1: GoTo E
Debug.Print jj.Fmt_Str_Repeat_Lv(mA, cLnItm, vbLf)
Exit Function
E: Gen_LetXXX = True
End Function

Function UsrPrf_ValidateLogin(pNmUsr$, pPwd$) As Boolean
Const cSub$ = "ValidateLogin"
Const cMsg$ = "Invalid User Id / Password"
If pPwd = "" Or pNmUsr = "" Then ss.A 1, cMsg, eUsrInfo: GoTo E
On Error GoTo R
Dim mRs As DAO.Recordset: Set mRs = CurrentDb.OpenRecordset("select * from tblUsr where NmUsr='" & pNmUsr & cQSng)
With mRs
    If .AbsolutePosition = -1 Then ss.A 2, cMsg, eUsrInfo: GoTo E
    If Not !Enabled Then .Close: ss.A 3, "User is disabled", eUsrInfo: GoTo E
    If pPwd <> !Password.Value Then ss.A 4, cMsg, eUsrInfo: GoTo E
    '=====Login is OK
    If UsrPrf_zzGetUsrPrf_ByRs(mRs) Then .Close: ss.A 5: GoTo E
    If UsrPrf_zzSetLoginToReg(!Usr.Value, True) Then ss.A 6: GoTo E
    .Edit
    !LoginCnt = !LoginCnt + 1
    !LasLoginDte = Now
    .Update
    .Close
End With
Exit Function
R: ss.R
E: UsrPrf_ValidateLogin = True: ss.B cSub, cMod, "pNmUsr,pPwd", pNmUsr, pPwd
    'If zzSetLoginToReg(0, False) Then ss.A 7: GoTo E
X: jj.Cls_Rs mRs
End Function
Function UsrPrf_zzChkLogin() As Boolean
Const cSub$ = "zzChkLogin"
Dim mApp$: mApp = jj.SysCfg_App
Dim mA$: mA = GetSetting(mApp, "UsrPrf", "HasLogin"): If mA <> "True" Then UsrPrf_zzLoginAgain mApp: Exit Function
Dim mAccessTim As Date: mAccessTim = GetSetting(mApp, "UsrPrf", "AccessTim")
If DateDiff("h", mAccessTim, Now()) > 1 Then UsrPrf_zzLoginAgain (mApp): Exit Function
If jj.UsrPrf_Usr <= 0 Then
    Dim mUsr%: mUsr% = GetSetting(mApp, "UsrPrf", "Usr")
    If mUsr <= 0 Then ss.xx 1, cSub, cMod, eQuit, "Cannot get user id": Application.Quit
    If UsrPrf_zzGetUsrPrf_ByUsr(mUsr) Then ss.xx 2, cSub, cMod, eQuit, "See PrvMsg": Application.Quit
End If
VBA.Interaction.SaveSetting mApp, "UsrPrf", "AccessTim", Now
End Function
Private Function UsrPrf_zzGetUsrPrf_ByRs(pRs As DAO.Recordset) As Boolean
Const cSub$ = "zzGetUsrPrf_ByRs"
On Error GoTo R
With pRs
    On Error Resume Next
    xUsrPrf.Brand = Nz(!Brand, 0)
    xUsrPrf.Env = Nz(!Env, 0)
    xUsrPrf.Dpt = Nz(!Dpt, 0)
    xUsrPrf.Usr = Nz(!Usr, 0)
    xUsrPrf.Lvl = Nz(!Lvl, "")
    xUsrPrf.Fy = Nz(!Fy, jj.Cv_Dte2FyNo)
End With
Exit Function
R: ss.R
E: UsrPrf_zzGetUsrPrf_ByRs = True: ss.B cSub, cMod
End Function
Private Function UsrPrf_zzGetUsrPrf_ByUsr(pUsr%) As Boolean
Const cSub$ = "zzGetUsrPrf_ByUsr"
On Error GoTo R
jj.Crt_Tbl_FmLnkNmt jj.Sffn_Dta, "tblUsr"
Dim mRs As DAO.Recordset: Set mRs = CurrentDb.OpenRecordset("Select * from tblUsr where Usr=" & pUsr)
If UsrPrf_zzGetUsrPrf_ByRs(mRs) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: UsrPrf_zzGetUsrPrf_ByUsr = True: ss.B cSub, cMod, "pUsr", pUsr
End Function
Function UsrPrf_zzLoginAgain(pApp$) As Boolean
Const cSub$ = "zzLoginAgain"
'Aim: Must login success (check against tblUsr), otherwise, quit.  If success, Reg: Usr & "AccessTim is set & xUsrPrf will be set
If jj.SysCfg_IsNoLogin Then
    VBA.Interaction.SaveSetting pApp, "UsrPrf", "AccessTim", Now
    On Error GoTo R
    If jj.Crt_Tbl_FmLnkNmt(jj.Sffn_Dta, "tblUsr") Then ss.xx 1, cSub, cMod, eQuit, "Cannot link tblUsr": Application.Quit
    Dim mRs As DAO.Recordset:   Set mRs = CurrentDb.OpenRecordset("select * from tblUsr")    ' Assume user is not one record
    If mRs.AbsolutePosition <> -1 Then
        If UsrPrf_zzGetUsrPrf_ByRs(mRs) Then mRs.Close: ss.xx 2, cSub, cMod, eQuit: Application.Quit
        VBA.Interaction.SaveSetting pApp, "UsrPrf", "Usr", xUsrPrf.Usr
        mRs.Close
        Exit Function
    End If
    With mRs
        .AddNew
        !NmUsr = "NoLogin"
        !Password = "password"
        !UsrLvl = "T"
        .Update
        .Close
    End With
    Set mRs = CurrentDb.TableDefs("tblUsr").OpenRecordset
    If UsrPrf_zzGetUsrPrf_ByRs(mRs) Then mRs.Close: ss.xx 2, cSub, cMod, eQuit: Application.Quit
    VBA.Interaction.SaveSetting pApp, "UsrPrf", "Usr", xUsrPrf.Usr
    mRs.Close
    Exit Function
End If
If jj.Opn_Frm("frmLoginAgain", , True) Then ss.xx 3, cSub, cMod, eQuit: Application.Quit
VBA.Interaction.SaveSetting pApp, "UsrPrf", "AccessTim", Now
VBA.Interaction.SaveSetting pApp, "UsrPrf", "Usr", xUsrPrf.Usr

Exit Function
R: ss.R
E: ss.B cSub, cMod
X: UsrPrf_zzSetLoginToReg 0, False
Application.Quit
End Function
Function UsrPrf_zzSetLoginToReg(pUsr%, pLoginOk As Boolean) As Boolean
Dim mApp$: mApp = jj.SysCfg_App
VBA.Interaction.SaveSetting mApp, "UsrPrf", "HasLogin", pLoginOk
VBA.Interaction.SaveSetting mApp, "UsrPrf", "AccessTim", Now
VBA.Interaction.SaveSetting mApp, "UsrPrf", "Usr", pUsr
End Function
'------------------------
Private Function xx() As tUsrPrf
UsrPrf_zzChkLogin
xx = xUsrPrf
End Function
