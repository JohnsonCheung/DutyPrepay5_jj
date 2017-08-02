Attribute VB_Name = "xSdir"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xSdir"
Function Sdir_Doc$()
Static xDir$:
If xDir = "" Then
    Dim mHom$: mHom = jj.Sdir_Hom
    If Len(mHom) = 3 And Right(mHom, 2) = ":\" Then
        xDir = mHom & "Documents\"
    Else
        xDir = mHom & "..\Documents\"
    End If
    jj.Crt_Dir xDir
End If
Sdir_Doc = xDir
End Function
Function Sdir_EmptyXls$()
Static xDir$: If xDir = "" Then xDir = jj.Sdir_Wrk & "EmptyXls\"
Sdir_EmptyXls = xDir
End Function
Function Sdir_ExpPgm$()
Static xDir$: If xDir = "" Then xDir = jj.Sdir_Doc & "Pgm\": jj.Crt_Dir xDir
Sdir_ExpPgm = xDir
End Function
Function Sdir_Hom$()
Static xDir$
If xDir = "" Then
    Dim mFfn$: mFfn = jj.Sffn_This
    Dim mA$
    mA = "WorkingDir\PgmObj\jj.mda": If IsEnd(mFfn, mA) Then xDir = Left(mFfn, Len(mFfn) - Len(mA)): Sdir_Hom = xDir: Exit Function
    mA = "jj.xla":                   If IsEnd(mFfn, mA) Then xDir = Left(mFfn, Len(mFfn) - Len(mA)): Sdir_Hom = xDir: Exit Function
    ss.A 1, "The core lib must named as jj.mda or jj.xla", eCritical: Stop: Application.Quit
End If
Sdir_Hom = xDir
End Function
Function Sdir_PgmObj$()
Sdir_PgmObj = jj.Sdir_Wrk & "PgmObj\"
End Function
Function Sdir_PgmLgc$()
Sdir_PgmLgc = jj.Sdir_Wrk & "PgmObj\Lgc\"
End Function
Function Sdir_Rpt$()
Static xDir$: If xDir = "" Then xDir = jj.Sdir_Hom & "Reports\": jj.Crt_Dir xDir
Sdir_Rpt = xDir
End Function
Function Sdir_RptSess$(pNmSess$)
Dim mDir$: mDir = jj.Sdir_Rpt & pNmSess & "\"
Static xDir$: If xDir <> mDir Then xDir = mDir: jj.Crt_Dir xDir
Sdir_RptSess = xDir
End Function
Function Sdir_Tmp$()
Static xDir$: If xDir = "" Then xDir = jj.SysCfg_DirTmp
Sdir_Tmp = xDir
End Function
Function Sdir_TmpApp$()
Static xDir$: If xDir = "" Then xDir = jj.Sdir_Tmp & jj.SysCfg_App & "\": jj.Crt_Dir xDir
Sdir_TmpApp = xDir
End Function
Function Sdir_TmpRqp$()
Static xDir$: If xDir = "" Then xDir = jj.Sdir_TmpApp & "Rqp\": jj.Crt_Dir xDir
Sdir_TmpRqp = xDir
End Function
Function Sdir_TmpLgc$()
Static xDirTmpLgc$: If xDirTmpLgc = "" Then xDirTmpLgc = jj.Sdir_Tmp & "Lgc\": jj.Crt_Dir xDirTmpLgc
Static xDir$: If xDir = "" Then xDir = xDirTmpLgc & jj.SysCfg_App & "\": jj.Crt_Dir xDir
Sdir_TmpLgc = xDir
End Function
Function Sdir_TmpRpt$(pNmRptSht$)
Dim mDir$: mDir = jj.Sdir_Tmp & pNmRptSht & "\"
Static xDir$: If xDir <> mDir Then xDir = mDir: jj.Crt_Dir xDir
Sdir_TmpRpt = xDir
End Function
Function Sdir_TmpSdir$(pSdir$)
Dim mDir$: mDir = jj.Sdir_Tmp & pSdir & "\"
Static xDir$: If xDir <> mDir Then xDir = mDir: jj.Crt_Dir xDir
Sdir_TmpSdir = xDir
End Function
Function Sdir_TmpSess$(pNmLgs$, pTrc&)
Dim mDir$: mDir = jj.Sdir_TmpLgc$ & pNmLgs & "\"
Static xDir1$: If xDir1 <> mDir Then xDir1 = mDir: jj.Crt_Dir xDir1
mDir = mDir & Format(pTrc, "00000000") & "\"
Static xDir2$: If xDir2 <> mDir Then xDir2 = mDir: jj.Crt_Dir xDir2
Sdir_TmpSess = xDir2
End Function
#If Tst Then
Function Sdir_TmpSess_Tst() As Boolean
Debug.Print jj.Sdir_TmpSess("MPS", 1)
End Function
#End If
Function Sdir_Tp$()
Sdir_Tp = jj.Sdir_Wrk & "Templates\"
End Function
Function Sdir_Wrk$()
Sdir_Wrk = jj.Sdir_Hom & "WorkingDir\"
End Function
Function Sdir_Wrk_YrWk$(pYr As Byte, pWk As Byte)
Sdir_Wrk_YrWk = jj.Sdir_Wrk & jj.ToStr_YrWk(pYr, pWk)
End Function
