Attribute VB_Name = "xXls"
Option Compare Database
Option Explicit
Const cMod$ = "xXls"
Function SetColOutLine(pRge As Range, pNCol As Byte, pNLvl As Byte) As Boolean
'Aim: use pRno:pCnoBeg-pCnoEnd to set column outline.
On Error GoTo E
If IsNothing(pRge) Then MsgBox "pRge is nothing": GoTo E
If pRge.Count > 1 Then MsgBox "pRge should not one cell.  pRge=" & pRge.Address: GoTo E
Dim iCno%, iRno&
Dim mWs As Worksheet: Set mWs = pRge.Parent
Dim mRge As Range
Dim iLvl As Byte
For iCno = pRge.Column + 1 To pRge.Column + pNCol - 1
    For iLvl = 2 To pNLvl
        iRno = pRge.Row + iLvl
        If IsEmpty(mWs.Cells(iRno, iCno).Value) Then
            Set mRge = mWs.Cells(1, iCno)
            Set mRge = mRge.EntireColumn
            mRge.OutlineLevel = iLvl
            GoTo NxtCol
        End If
    Next
    Set mRge = mWs.Cells(1, iCno)
    Set mRge = mRge.EntireColumn
    mRge.OutlineLevel = iLvl
NxtCol:
Next
Dim mAyRgeCno() As tRgeCno
For iLvl = 0 To 2
    iRno = iLvl + pRge.Row
    Set mRge = mWs.Range(mWs.Cells(iRno, pRge.Column), mWs.Cells(iRno, pRge.Column + pNCol - 1))
    If Fnd_AyRgeCno(mAyRgeCno, mRge) Then GoTo E
    Dim J%
    For J = 0 To UBound(mAyRgeCno)
        With mAyRgeCno(J)
            MgeRge mWs.Range(mWs.Cells(iRno, .Fm), mWs.Cells(iRno, .To))
        End With
    Next
Next
Exit Function
E: SetColOutLine = True
End Function
Function Fnd_AyRgeCno(oAyRgeCno() As tRgeCno, pRge As Range) As Boolean
'Aim: find {oAyRgeCno} by each 'block'.  One 'block' one element in {oAyRgeCno}.
'     a 'block' is pCol for of CnoFm & CnoTo having same value.
'     pRge must a one line of cells.
'     Note: empty cell block will not included.
Const cSub$ = "Fnd_AyRgeCno"
On Error GoTo R
If pRge.Count = 1 Then Exit Function
'-- Find mNCol
Dim mNCol As Byte: mNCol = pRge.Columns.Count
Dim iCno As Byte
Dim mN%: mN = 0
Dim mWs As Worksheet: Set mWs = pRge.Parent
Dim mRno&: mRno = pRge.Row
Dim mV, mVLas
Dim mCnoFm As Byte: mCnoFm = pRge.Column
mVLas = pRge.Cells(1, 1).Value

iCno = mCnoFm
Do
    If IsEmpty(mVLas) Then
        mCnoFm = iCno
        mVLas = mWs.Cells(mRno, iCno).Value
        GoTo Nxt
    End If
           
    If mVLas <> mWs.Cells(mRno, iCno).Value Then
        mVLas = mWs.Cells(mRno, iCno).Value
        ReDim Preserve oAyRgeCno(mN)
        With oAyRgeCno(mN)
            .Fm = mCnoFm
            .To = iCno - 1
        End With
        mN = mN + 1
        mCnoFm = iCno
    End If
Nxt:
    iCno = iCno + 1
Loop Until iCno > pRge.Column + mNCol - 1

If iCno > mCnoFm Then
    ReDim Preserve oAyRgeCno(mN)
    With oAyRgeCno(mN)
        .Fm = mCnoFm
        .To = iCno - 1
    End With
End If
Exit Function
R: ss.R
E: Fnd_AyRgeCno = True: ss.B cSub, cMod, "pRge", jj.ToStr_Rge(pRge)
End Function
Function SetColOutLineColr(pRge As Range, pNLvl As Byte) As Boolean
'Aim: pRge
On Error GoTo R
Const cSub$ = "SetColOutLineColr"
'--
Dim mRge As Range
Dim mWs As Worksheet: Set mWs = pRge.Worksheet
'-- Find mAyRgeRno(): Use first column & pNLvl & color in the cells to find which RgeRno of cells needs to set color
Dim mAyRgeRno() As tRgeRno, mN%
mN = 0
Dim iRno&
Dim mColrWhite&:  mColrWhite = 16777215
Dim mRnoFm&: mRnoFm = 0
For iRno = pRge.Row + pNLvl To pRge.Row + pRge.Rows.Count - 1
    Set mRge = mWs.Cells(iRno, pRge.Column)
    If mRge.Interior.Color = mColrWhite Then
        If mRnoFm <> 0 Then
            ReDim Preserve mAyRgeRno(mN)
            mAyRgeRno(mN).Fm = mRnoFm
            mAyRgeRno(mN).To = iRno - 1
            mN = mN + 1
            mRnoFm = 0
        End If
        GoTo Nxt
    End If
    If mRnoFm = 0 Then mRnoFm = iRno
Nxt:
Next
If mRnoFm <> 0 Then
    ReDim Preserve mAyRgeRno(mN)
    mAyRgeRno(mN).Fm = mRnoFm
    mAyRgeRno(mN).To = iRno - 1
End If
'-- Find mAyColr(1-pNLvl) by pRge downward pNLvl cells
ReDim mAyColr&(pNLvl)
Dim iLvl As Byte
For iLvl = 1 To pNLvl - 1
    Set mRge = pRge.Cells(iLvl, 1)
    mAyColr(iLvl) = mRge.Interior.Color
Next
'-- Loop all col and set color
Dim iCno As Byte
For iCno = pRge.Column To pRge.Column + pRge.Columns.Count - 1
    iLvl = mWs.Columns(iCno).OutlineLevel
    If iLvl < pNLvl Then
        Dim mColr&: mColr = mAyColr(iLvl)
        '--
        Set mRge = mWs.Range(mWs.Cells(pRge.Row + iLvl, iCno), mWs.Cells(pRge.Row + pNLvl - 1, iCno))
        mRge.MergeCells = True
        mRge.Interior.Color = mColr
        mRge.Borders(xlEdgeTop).LineStyle = xlLineStyleNone
        '
        Dim J As Byte
        For J = 0 To UBound(mAyRgeRno)
            Set mRge = mWs.Range(mWs.Cells(mAyRgeRno(J).Fm, iCno), mWs.Cells(mAyRgeRno(J).To, iCno))
            mRge.Interior.Color = mColr
        Next
    End If
Next
Exit Function
R: ss.R
E: SetColOutLineColr = True: ss.B cSub, cMod, "pRge", jj.ToStr_Rge(pRge)
End Function
Function MgeRge(pRge As Range) As Boolean
If IsNothing(pRge) Then Exit Function
If pRge.Count = 1 Then Exit Function
Dim mVal$: mVal = pRge.Cells(1, 1).Value
pRge.Value = Null
pRge.Cells(1, 1).Value = mVal
pRge.Merge
End Function
