VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".cSq"
Public Rno1&, Rno2&, Cno1 As Byte, Cno2 As Byte
Public Function GetRge(oRge As Range, pWs As Worksheet) As Boolean
GetRge = True
If 1 > Cno1 Or Cno1 > 255 Then MsgBox "cSq.GetRge: invalid Cno1 (Cno1=" & Cno1 & ")": Exit Function
If Cno1 > Cno2 Or Cno2 > 255 Then MsgBox "cSq.GetRge: invalid Cno2 (Cno1=" & Cno1 & ", Cno2=" & Cno2 & ")": Exit Function
If 1 > Rno1 Or Rno1 > 65536 Then MsgBox "cSq.GetRge: invalid Rno1 (Rno1=" & Rno1 & ")": Exit Function
If Rno1 > Rno2 Or Rno2 > 65536 Then MsgBox "cSq.GetRge: invalid Rno2 (Rno1=" & Rno1 & ", Rno2=" & Rno2 & ")": Exit Function
With pWs
    Set oRge = .Range(.Cells(Rno1, Cno1), .Cells(Rno2, Cno2))
End With
GetRge = False
End Function
Public Function ToStr$()
ToStr$ = "(" & Rno1 & cComma & Cno1 & "),(" & Rno2 & cComma & Cno2 & ")"
End Function
