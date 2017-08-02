Attribute VB_Name = "xIns"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xIns"
Function Ins_VCells(pWs As Worksheet, pLoCol$, pRnoBeg&, pNRow&) As Boolean
Dim mAyCol$(), J As Byte
mAyCol = Split(pLoCol, cComma)
With pWs
    For J = 0 To UBound(mAyCol)
        .Range(mAyCol(J) & pRnoBeg & ":" & mAyCol(J) & pRnoBeg + pNRow - 1).Insert xlShiftToRight
    Next
End With
End Function


