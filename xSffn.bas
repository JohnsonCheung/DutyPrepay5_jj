Attribute VB_Name = "xSffn"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xSffn"
Function Sffn_Cur$()
Static xFfn$
If xFfn = "" Then
    If jj.IsXls Then
        xFfn = ThisWorkbook.Name
     ElseIf jj.IsAcs Then
        xFfn = Fct.Nam_FilNam(CurrentDb.Name)
    End If
End If
Sffn_Cur = xFfn
End Function
Function Sffn_DbLog$()
Sffn_DbLog = jj.Sdir_PgmObj & "Log.Mdb"
End Function
Function Sffn_SessDta$(pTrc&)
Sffn_SessDta = Fct.CurMdbDir & Format(pTrc, "00000000") & "\" & Fct.CurMdbNam & "_Dta.mdb"
End Function
Function Sffn_Dta$()
Static xSffn$: If xSffn = "" Then xSffn = jj.Sdir_Wrk & jj.SysCfg_App & "_Data.Mdb"
Sffn_Dta = xSffn
End Function
Function Sffn_Rpt$(pNmRptSht$, Optional pNmSess$ = "", Optional pTimStampOpt As eTimStampOpt = eDte, Optional pInRptDir As Boolean = False, Optional pExt$ = ".xls")
Dim mTimStamp$
Select Case pTimStampOpt
Case eYr:  mTimStamp = Format(Date, "_yyyy")
Case eMth: mTimStamp = Format(Date, "_yyyy_mm")
Case eWk:  mTimStamp = Year(Date) & "W" & Format(Format(Date, "ww"), "00")
Case eDte: mTimStamp = Format(Date, "_yyyy_mm_dd")
Case eMin: mTimStamp = Format(Now, "_yyyy_mm_dd@HHMM")
End Select

If pInRptDir Then
    Static xNmRptSht$: If xNmRptSht <> pNmRptSht Then xNmRptSht = pNmRptSht: jj.Crt_Dir jj.Sdir_Rpt & pNmRptSht & "\"
    Sffn_Rpt = jj.Sdir_Rpt & pNmRptSht & "\" & jj.Fmt_Str("{0}{1}{2}", pNmSess, mTimStamp, pExt)
Else
    Dim mNmSess$: If pNmSess <> "" Then mNmSess = jj.Q_S(pNmSess, "()")
    Sffn_Rpt = jj.Sdir_Rpt & jj.Fmt_Str("{0}{1}_{2}{3}", pNmRptSht, mNmSess, mTimStamp, pExt)
End If
End Function
Function Sffn_This$()
Const cSub$ = "This"
Static A$
If A$ = "" Then
    Dim mPrj As VBProject: If jj.Fnd_Prj(mPrj, cLib) Then ss.xx 1, cSub, cMod, eCritical: Stop: Application.Quit
    A = mPrj.Filename
End If
Sffn_This = A$
End Function
#If Tst Then
Function Sffn_This_Tst() As Boolean
Debug.Print jj.Sffn_This
End Function
#End If
Function Sffn_TmpAppUsrMdb$()
Static xSffn$: If xSffn = "" Then xSffn = jj.Sdir_TmpApp() & jj.Fmt_Str("tmp{0}_{1}.mdb", jj.SysCfg_App, jj.UsrPrf_Usr)
Sffn_TmpAppUsrMdb = xSffn
End Function
Function Sffn_SessTp$(pNmLgs$, pNmTp$, pTrc&)
Sffn_SessTp = jj.Sdir_TmpSess(pNmLgs, pTrc) & jj.Fmt_Str("tmp{0}_{1}_{2}.mdb", jj.SysCfg_App, jj.UsrPrf_Usr, pNmTp$)
End Function
Function Sffn_Tp$(pNmRptSht$, Optional pNmSess$ = "", Optional pExt$ = ".xls")
Dim mNmSess$: If pNmSess <> "" Then mNmSess = "_" & pNmSess
Sffn_Tp = jj.Sdir_Tp & jj.Fmt_Str("Template_{0}{1}" & pExt, pNmRptSht, mNmSess)
End Function

