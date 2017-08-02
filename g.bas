Attribute VB_Name = "g"
#Const Tst = True
#Const PDF = False
Option Compare Text
Option Explicit
Public Const cLib$ = "jj"
Const cMod$ = cLib & ".g"
'Private xUsrPrf As tUsrPrf
Public gIsBch As Boolean, gSilent As Boolean, gNoLog As Boolean
Public Const cColrNo& = 16777215
Public Const cColrCmd& = 16776960
Public Const cColrAdd& = 128 'Brown
Public Const cColrDlt& = 255 'Red
Public Const cRnoDta& = 5
Public Const cRnoNmFld& = 4
Public Const cQSng$ = "'"
Public Const cQDbl$ = """"
Public Const cSemi$ = ";"
Public Const cComma$ = ","
Public Const cCommaSpc$ = ", "
'------------------------
Private x_IsCnl As Boolean
Private x_DteSel As Variant
Private x_Xls As Excel.Application
Private x_Acs As Access.Application
Private x_Wrd As Word.Application
Private x_Ppt As PowerPoint.Application
Function gPpt() As PowerPoint.Application
If jj.IsNothing(x_Ppt) Then
    Set x_Ppt = New PowerPoint.Application
Else
    On Error GoTo R
    Dim mA$: mA = x_Ppt.Name
End If
Set gPpt = x_Ppt
Exit Function
R: ss.R
    Set x_Ppt = New PowerPoint.Application
    Set gPpt = x_Ppt
End Function
Function gAcs() As Access.Application
If jj.IsNothing(x_Acs) Then
    Set x_Acs = New Access.Application
Else
    On Error GoTo R
    Dim mA$: mA = x_Acs.Name
End If
Set gAcs = x_Acs
Exit Function
R: ss.R
    Set x_Acs = New Access.Application
    Set gAcs = x_Acs
End Function
Function gWrd() As Word.Application
If jj.IsNothing(x_Wrd) Then
    Set x_Wrd = New Word.Application
Else
    On Error GoTo R
    Dim mA$: mA = x_Wrd.Name
End If
Set gWrd = x_Wrd
Exit Function
R: ss.R
    Set x_Wrd = New Word.Application
    Set gWrd = x_Wrd
End Function
Function gXls() As Excel.Application
If jj.IsNothing(x_Xls) Then
    Set x_Xls = New Excel.Application
Else
    On Error GoTo R
    Dim mA$: mA = x_Xls.Name
End If
Set gXls = x_Xls
Exit Function
R: ss.R
    Set x_Xls = New Excel.Application
    Set gXls = x_Xls
End Function
Function gDbEng() As DAO.DBEngine
Static x_DbEng As DAO.DBEngine
If jj.IsAcs Then Set gDbEng = Application.DBEngine: Exit Function
If jj.IsNothing(x_DbEng) Then Set x_DbEng = New DAO.DBEngine
Set gDbEng = x_DbEng
End Function
'------------------------
Property Get gDteSel()
gDteSel = x_DteSel
End Property
Property Let gDteSel(pDte)
x_DteSel = pDte
End Property
Public Property Get gFmtCnt$(): gFmtCnt$ = "#,##0":               End Property
Public Property Get gFmtCur$(): gFmtCur$ = "#,##0.00;(#,##0.00)": End Property
Public Property Get gFmtDte$(): gFmtDte$ = "yyyy/mm/dd":          End Property
Function gFso() As FileSystemObject
Static xFso As FileSystemObject
If jj.IsNothing(xFso) Then Set xFso = New FileSystemObject
Set gFso = xFso
End Function
Property Get gIsCnl() As Boolean
gIsCnl = x_IsCnl
End Property
Property Let gIsCnl(pIsCnl As Boolean)
x_IsCnl = pIsCnl
End Property
'Function gOL() As Outlook.Application
'Static xOL As Outlook.Application: If jj.IsNothing(xOL) Then Set xOL = New Outlook.Application
'Set gOL = xOL
'End Function
'#If PDF Then
'Function gPDF() As PDFCreator.clsPDFCreator
'Static xPDF As PDFCreator.clsPDFCreator: If jj.IsNothing(xPDF) Then Set xPDF = New PDFCreator.clsPDFCreator: xPDF.cStart "/NoProcessingAtStartup"
'Set gPDF = xPDF
'End Function
'#End If
'------
