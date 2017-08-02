Attribute VB_Name = "xRbr"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".Rbr"
Function Rbr_ByRs(pRs As DAO.Recordset, pStart As Byte, pStp As Byte) As Boolean
Dim I&: I = pStart
With pRs
    While Not .EOF
        .Edit
        .Fields(0).Value = I: I = I + pStp
        .Update
        .MoveNext
    Wend
    .Close
End With
End Function
Function Rbr_ByTbl(pTbl$, pNmfld_ToRbr$, Optional pStart As Byte = 1, Optional pStp As Byte = 1) As Boolean
Dim mRs As DAO.Recordset: Set mRs = CurrentDb.OpenRecordset("Select {0} from {1} order by {0}", pNmfld_ToRbr, pTbl)
Rbr_ByTbl = Rbr_ByRs(mRs, pStart, pStp)
End Function

