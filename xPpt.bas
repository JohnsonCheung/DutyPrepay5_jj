Attribute VB_Name = "xPpt"
Option Compare Database
Option Explicit
Const cMod$ = cLib & ".xxPpt"
Function GetShape_ByStr(pPpt As Presentation, pS$) As PowerPoint.Shape
Dim iSlide As Slide
Dim iShape As PowerPoint.Shape
Dim iTxtFme As TextFrame
Dim iTxtRge As TextRange
For Each iSlide In pPpt.Slides
    For Each iShape In iSlide.Shapes
        If iShape.Type = msoTextBox Then
            Set iTxtRge = iShape.TextFrame.TextRange.Find(pS)
            If TypeName(iTxtRge) <> "Nothing" Then Set GetShape_ByStr = iShape: Exit Function
        End If
    Next
Next
End Function
Function RplShape_ByFx(pShape As PowerPoint.Shape, pFx$) As Boolean
On Error GoTo E
Dim mObj As Object: Set mObj = pShape.Parent
If TypeName(mObj) <> "Slide" Then MsgBox "pShape.Parent is not Slide": GoTo E
Dim mSlide As Slide: Set mSlide = pShape.Parent
mSlide.Select
pShape.Delete
mSlide.Shapes.AddOLEObject Left:=20, Top:=20, Width:=670, Height:=500, Filename:=pFx
Exit Function
E: RplShape_ByFx = True
End Function
Public Function Opn_Fp(pFp$) As Boolean
On Error GoTo R
If Not IsFfn(pFp, pSilient:=False) Then GoTo E
Dim mPpt As New PowerPoint.Application
mPpt.Visible = True
mPpt.Presentations.Open pFp, Untitled:=msoTrue
Exit Function
R: ss.R
E: Opn_Fp = True
End Function
Public Function Opn_Ppt(ByRef oPpt As PowerPoint.Presentation, pFp$, Optional pIsInNewPpt As Boolean = False) As Boolean
Const cSub$ = "Opn_Ppt"
On Error GoTo R
Dim mPpt As PowerPoint.Application
If pIsInNewPpt Then
    Set mPpt = New PowerPoint.Application
Else
    Set mPpt = gPpt
End If
mPpt.Visible = True
mPpt.WindowState = ppWindowMinimized
Set oPpt = mPpt.Presentations.Open(pFp, Untitled:=msoTrue)
Exit Function
R: ss.R
E: Opn_Ppt = True: ss.B cSub, cMod, "pFp,pIsInNewPpt", pFp, pIsInNewPpt
End Function
Function Sav_Ppt(pPpt As PowerPoint.Presentation) As Boolean
On Error GoTo E
pPpt.Save
Exit Function
E: Sav_Ppt = True
End Function
Function SavAs_Ppt(pPpt As PowerPoint.Presentation, pFp$) As Boolean
On Error GoTo E
pPpt.SaveAs pFp
Exit Function
E: SavAs_Ppt = True
End Function


