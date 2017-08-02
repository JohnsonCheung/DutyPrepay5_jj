Attribute VB_Name = "xEdt"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".Edt"
Function Edt_QrySql_ByUEDIT32(Optional pPfx$ = "qry") As Boolean
Const cSub$ = "Edt_QrySql_ByUEDIT32"
Dim mDir$: mDir = jj.Sdir_TmpSdir("Sql")
Dim L%: L = Len(pPfx)
Dim mA$
Dim iQry As QueryDef: For Each iQry In CurrentDb.QueryDefs
    If Left(iQry.Name, L) = pPfx Then
        Dim mFno As Byte: If jj.Opn_Fil_ForOutput(mFno, mDir & iQry.Name) Then ss.A 1: GoTo E
        Open mDir & iQry.Name & ".sql" For Output As #1
        Print #mFno, iQry.Sql
        Close #mFno
        mA = mA & jj.Fmt_Str("""{0}{1}.sql"" ", mDir, iQry.Name)
    End If
Next
If mA = "" Then MsgBox "No queries with given prefix [" & pPfx & "]": Exit Function
Edt_Txt mA
Exit Function
R: ss.R
E: Edt_QrySql_ByUEDIT32 = True: ss.B cSub, cMod, "pPfx", pPfx
End Function
Function Edt_Txt(pFfnTxt$) As Boolean
Shell jj.Fmt_Str("NotePad ""{0}""", pFfnTxt$), vbMaximizedFocus
End Function
