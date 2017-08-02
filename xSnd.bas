Attribute VB_Name = "xSnd"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xSnd"
Type tSmtpCfg
    SmtpServer As String
    SmtpUseSsl As Boolean
    SmtpServerPort As Integer
    SmtpAuthenticate As cdo.CdoProtocolsAuthentication
    SendUser As String
    SendPassword As String
End Type
Public xSmtpCfg As tSmtpCfg
Function Snd_LogMdb() As Boolean
Const cSub$ = "Snd_LogMdb"
If Not Fct.Start("Send the Log.Mdb to johnsoncheunghk@yahoo.com.hk") Then Exit Function
ss.ClsDbLog
Dim mFfnTmpLog$: mFfnTmpLog = jj.Sdir_Tmp & "Log.mdb_"
If jj.Cpy_Fil(jj.Sffn_DbLog, mFfnTmpLog) Then ss.A 1: GoTo E
Dim mBody$: mBody = "SdirHom=" & jj.Sdir_Hom & vbLf & "SffnDbLog=" & jj.Sffn_DbLog
Snd_LogMdb = jj.Snd_Mail("Snd.LogMdb", "johnsoncheunghk06@yahoo.com.hk", "", "LogMdb @" & Now(), mBody, mFfnTmpLog)
jj.Dlt_Fil mFfnTmpLog
Exit Function
R: ss.R
E: Snd_LogMdb = True: ss.B cSub, cMod
End Function
Function Snd_LogMdb_Tst() As Boolean
Debug.Print jj.Snd_LogMdb
End Function
Function Snd_Mail(pFm$, pTo$, pCC$, pSubj$, pBody$, Optional pFfnAttach$ = "", Optional pIsBySMTP As Boolean = False) As Boolean
Const cSub$ = "Snd_Mail"
If pFfnAttach <> "" Then If Not jj.IsFfn(pFfnAttach) Then ss.A 1: GoTo E
If pIsBySMTP Then GoTo SMTP
'By Outlook
On Error GoTo R
'Dim mMail_OL As Outlook.MailItem: Set mMail_OL = gOL.CreateItem(olMailItem)
'With mMail_OL
'    .To = pTo
'    '.SenderEmailAddress = pFm      .SenderEmailAddress is read only
'    If pFfnAttach <> "" Then .Attachments.Add pFfnAttach, , , Fct.Nam_FilNam(pFfnAttach)
'    .CC = pCC
'    .Subject = pSubj
'    .Body = pBody
'    .Send
'End With
'Exit Function
SMTP:
Dim mMail_CDO As New cdo.Message
With mMail_CDO
    .To = pTo
    .CC = pCC
    .From = pFm
    .Subject = pSubj
    .TextBody = pBody
    With .Configuration.Fields
        Dim mSmtpCfg As tSmtpCfg
        If xSmtpCfg.SmtpServer = "" Then
            With mSmtpCfg
                .SmtpServer = "127.0.0.1"
                .SmtpServerPort = 25
                .SmtpAuthenticate = cdo.CdoProtocolsAuthentication.cdoAnonymous
                .SmtpUseSsl = False
                .SendUser = ""
                .SendPassword = ""
            End With
        Else
            mSmtpCfg = xSmtpCfg
        End If
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing").Value = cdo.CdoSendUsing.cdoSendUsingPort
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = mSmtpCfg.SmtpServer ' "127.0.0.1" ' "localhost" ' "smtp.YourServer.com" ' Or "mail.server.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = mSmtpCfg.SmtpServerPort
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = mSmtpCfg.SmtpUseSsl
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = mSmtpCfg.SendPassword
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = mSmtpCfg.SendUser
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = mSmtpCfg.SmtpAuthenticate
        .Update
    End With
    '.Configuration.Fields.Append "SendUsing", adInteger, , , CDO.CdoSendUsing.cdoSendUsingPort
    If pFfnAttach <> "" Then .AddAttachment "file://" & pFfnAttach
    .Send
End With
Exit Function
R: ss.R
E: Snd_Mail = True: ss.B cSub, cMod, "pFm,pTo,pCC,pSubj,pBody,pFfnAttach", pFm, pTo, pCC, pSubj, pBody, pFfnAttach
End Function
Function Snd_Mail_ByYahoo(pFm$, pTo$, pCC$, pSubj$, pBody$, Optional pFfnAttach$ = "", Optional pIsBySMTP As Boolean = False) As Boolean
Const cSub$ = "Snd_Mail_ByYahoo"
If pFfnAttach <> "" Then If Not jj.IsFfn(pFfnAttach) Then ss.A 1: GoTo E
If pIsBySMTP Then GoTo SMTP
'By Outlook
On Error GoTo R
'Dim mMail_OL As Outlook.MailItem: Set mMail_OL = gOL.CreateItem(olMailItem)
'With mMail_OL
'    .To = pTo
'    '.SenderEmailAddress = pFm      .SenderEmailAddress is read only
'    If pFfnAttach <> "" Then .Attachments.Add pFfnAttach, , , Fct.Nam_FilNam(pFfnAttach)
'    .CC = pCC
'    .Subject = pSubj
'    .Body = pBody
'    .Send
'End With
'Exit Function
SMTP:
Dim mMail_CDO As New cdo.Message
With mMail_CDO
    .To = pTo
    .CC = pCC
    .From = pFm
    .Subject = pSubj
    .TextBody = pBody
    With .Configuration.Fields
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.mail.yahoo.com.hk" ' "localhost" ' "smtp.YourServer.com" ' Or "mail.server.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing").Value = cdo.CdoSendUsing.cdoSendUsingPort
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "ritachan"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "johnsoncheunghk06@yahoo.com.hk"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
        .Update
    End With
    '.Configuration.Fields.Append "SendUsing", adInteger, , , CDO.CdoSendUsing.cdoSendUsingPort
    If pFfnAttach <> "" Then .AddAttachment "file://" & pFfnAttach
    .Send
End With
Exit Function
R: ss.R
E: Snd_Mail_ByYahoo = True: ss.B cSub, cMod, "pFm,pTo,pCC,pSubj,pBody,pFfnAttach", pFm, pTo, pCC, pSubj, pBody, pFfnAttach
End Function
Function Snd_Mail_ByEnv(pEnv As tEnv, Optional pIsBySMTP As Boolean = True) As Boolean
With pEnv
    Snd_Mail_ByEnv = jj.Snd_Mail(.Fm, .To, .CC, .Subj, .Body, .Ffn, pIsBySMTP)
End With
End Function
Function Snd_Mail_Sample() As Boolean
'Aim: Sending a text email using authentication against a remote SMTP server
'     Sample code submitted by Clint Baldwin, +1 (918) 671 3429
Const cdoSendUsingPickup = 1
Const cdoSendUsingPort = 2
Const cdoAnonymous = 0
' Use basic (clear-text) authentication.
Const cdoBasic = 1
' Use NTLM authentication
Const cdoNTLM = 2 'NTLM
Dim objMEssage As cdo.Message
' Create the message object.
Set objMEssage = CreateObject("CDO.Message")
'Set the from address this would be your email address.
objMEssage.From = """Your Name""<Youremail@YourDomain.com>"
' Set the TO Address separate multiple address with a comma
objMEssage.To = "SomeEmail@YourDomain.com"
' Set the Subject.
objMEssage.Subject = "An Email From Active Call Center."
' Now for the Message Options Part.
' Use standared text for the body.
objMEssage.TextBody = _
"This is some sample message text.." & _
vbCrLf & _
"It was sent using SMTP authentication."

' Or you could use HTML as:
' objMessage.HTMLBody = strHTML

' ATTACHMENT : Add an attachment Can be any valid url'
'objMEssage.AddAttachment ("file://C:\Program Files\Active Call Center\Examples\Goodbye.wav")

' This section provides the configuration information for the SMTP server.
' Specifie the method used to send messages.
objMEssage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = _
cdoSendUsingPort

' The name (DNS) or IP address of the machine
' hosting the SMTP service through which
' messages are to be sent.
objMEssage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost" ' "smtp.YourServer.com" ' Or "mail.server.com"

' Specify the authentication mechanism
' to use.
objMEssage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = _
cdoBasic

' The username for authenticating to an SMTP server using basic (clear-text) authentication
objMEssage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusername") = _
"YourLogin@YourDomain.com"

' The password used to authenticate
' to an SMTP server using authentication
objMEssage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendpassword") = _
"Password"

' The port on which the SMTP service
' specified by the smtpserver field is
' listening for connections (typically 25)
objMEssage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = _
25

'Use SSL for the connection (False or True)
objMEssage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = _
False

' Set the number of seconds to wait for a valid socket to be established with the SMTP service before timing out.
objMEssage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = _
60

' Update configuration
objMEssage.Configuration.Fields.Update

' Use to show the message.
' MsgBox objMessage.GetStream.ReadText

' Send the message.
objMEssage.Send
End Function
#If Tst Then
Function Snd_Mail_Tst() As Boolean
Debug.Print jj.Snd_Mail("Snd.Mail_Tst", "johnsoncheunghk06@yahoo.com.hk", "jcheung@movadogroup.com", "Snd.Mai_Tst ....", "This is a testing", "c:\a.txt")
Dim J As Byte
For J = 1 To 2
    Debug.Print jj.Snd_Mail("johnsoncheunghk06@yahoo.com.hk", "johnsoncheunghk06@yahoo.com.hk", "johnsoncheunghk06@yahoo.com.hk", "subj -- Test", "body- This is test", "c:\a.txt", False)
Next
End Function
#End If
Function Snd_Tbl_ToMdb(pLikNmt$, pFbTar$, Optional pCrtIfNotExist As Boolean = False) As Boolean
'Aim: Currentdb db's {pLikNmt} tables to {pFbTar}
Const cSub$ = "Snd_Tbl_ToMdb"
If VBA.Dir(pFbTar) = "" Then
    If Not pCrtIfNotExist Then ss.A 1, "Given {pFbTar} not exist": GoTo E
    If jj.Crt_Fb(pFbTar) Then ss.A 2: GoTo E
End If
Dim mAnt$(): If jj.Fnd_Ant_ByLik(mAnt, pLikNmt) Then ss.A 4: GoTo E
Dim J%
For J = 0 To jj.Siz_Ay(mAnt) - 1
    Dim mSql$: mSql = jj.Fmt_Str("Select * into {0} in '{1}' from {0}", mAnt(J), pFbTar)
    If jj.Run_Sql(mSql) Then ss.A 5: GoTo E
Next
Exit Function
R: ss.R
E: Snd_Tbl_ToMdb = True: ss.B cSub, cMod, "pLikNmt,pFbTar,pCrtIfNotExit", pLikNmt, pFbTar, pCrtIfNotExist
End Function
#If Tst Then
Function Snd_Tbl_ToMdb_Tst() As Boolean
If jj.Snd_Tbl_ToMdb("tmp*", "c:\aa.mdb", True) Then Stop
End Function
#End If

