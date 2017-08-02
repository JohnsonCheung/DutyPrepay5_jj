Attribute VB_Name = "xLst"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".Lst"
Function Lst_CmdTxt(pWb As Workbook, Optional pFno As Byte = 0) As Boolean
Dim iWs As Worksheet, iQt As QueryTable, iPt As PivotTable
For Each iWs In pWb.Worksheets
    If iWs.PivotTables.Count > 0 Then
        jj.Prt_Ln pFno, Fct.UnderlineStr("Worksheet " & iWs.Name & " (PivotTables)", "-")
        For Each iPt In iWs.PivotTables
            jj.Prt_Ln pFno, jj.ToStr_Pt(iPt)
        Next
        jj.Prt_Ln pFno
    End If
    If iWs.QueryTables.Count > 0 Then
        jj.Prt_Ln pFno, Fct.UnderlineStr("Worksheet " & iWs.Name & " (QueryTables)", "-")
        For Each iQt In iWs.QueryTables
            jj.Prt_Ln pFno, jj.ToStr_Qt(iQt)
        Next
        jj.Prt_Ln pFno
    End If
Next
End Function
Function Lst_QryList(pNmqPfx$, Optional pSQL_SubString$ = "") As Boolean
Dim L%: L = Len(pNmqPfx)
Dim iQry As QueryDef: For Each iQry In CurrentDb.QueryDefs
    If Left(iQry.Name, L) = pNmqPfx Then If InStr(iQry.Sql, pSQL_SubString) > 0 Then Debug.Print jj.ToStr_TypQry(iQry.Type), iQry.Name
Next
End Function
Function Lst_QryPrm_ByPfx(pNmqPfx$, Optional pFno As Byte = 0) As Boolean
Dim L%: L = Len(pNmqPfx)
Dim iQry As QueryDef: For Each iQry In CurrentDb.QueryDefs
    If Left(iQry.Name, L) = pNmqPfx Then
        If iQry.Parameters.Count > 0 Then
            jj.Prt_Str pFno, iQry.Name & "-----(Param)------>"
            Dim iPrm As DAO.Parameter
            For Each iPrm In iQry.Parameters
                jj.Prt_Str pFno, iPrm.Name
            Next
            jj.Prt_Ln pFno
        End If
    End If
Next
End Function

