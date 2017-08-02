Attribute VB_Name = "xToSql"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xToSql"
Function ToSql_Dlt$(pNmt$, pLExpr$)
ToSql_Dlt = jj.Fmt_Str("Delete from {0}{1}", pNmt, jj.Cv_Str(pLExpr, " Where "))
End Function
Function ToSql_Ins$(pNmt$, pLst$, pVal$)
ToSql_Ins = jj.Fmt_Str("Insert into {0} ({1}) values ({2})", jj.Q_SqBkt(pNmt), pLst, pVal)
End Function
Function ToSql_Sel$(pNmt$, pSel$, Optional pLExpr$ = "")
ToSql_Sel = jj.Fmt_Str("Select {0} from {1}{2}", pSel, jj.Q_SqBkt(pNmt), jj.Cv_Where(pLExpr))
End Function
Function ToSql_Upd$(pNmt$, pLoAsg$, pLExpr$)
ToSql_Upd = jj.Fmt_Str("Update {0} Set {1}{2}", jj.Q_SqBkt(pNmt), pLoAsg, jj.Cv_Where(pLExpr))
End Function
Function ToSql_Nmq$(pNmq$)
Const cSub$ = "ToSql_Nmq"
On Error GoTo R
ToSql_Nmq = CurrentDb.QueryDefs(pNmq).Sql
Exit Function
R: ss.R
E: ss.B cSub, cMod, "pNmq", pNmq
ToSql_Nmq = "Err: jj.ToSql_Nmq('" & pNmq & "').  Msg=" & Err.Description
End Function
#If Tst Then
Function ToSql_Nmq_Tst() As Boolean
Debug.Print jj.ToSql_Nmq("qryMPS_1")
End Function
#End If
