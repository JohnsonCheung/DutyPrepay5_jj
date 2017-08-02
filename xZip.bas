Attribute VB_Name = "xZip"
#Const Tst = True
Option Compare Text
Option Base 0
Option Explicit
Const cMod$ = cLib & ".xZip"
Function Zip_By7Z(pFfn$, Optional pOvrWrt As Boolean = False) As Boolean
'Aim: Zip {pFfn} to file in same directory {pFfn} extension cannot be .zip
Const cSub$ = "Zip_By7Z"
Const cPgm$ = "C:\Program Files\7-zip\7z.exe"
If Right(pFfn, 4) = ".zip" Then ss.A 1, "Cannot end with .zip": GoTo E
If VBA.Dir(pFfn) = "" Then ss.A 2, "Given file does not exist": GoTo E
Dim mCmd$, mFfnZip$
mFfnZip$ = jj.Cut_Ext(pFfn) & ".zip"
If jj.Ovr_Wrt(mFfnZip, pOvrWrt) Then ss.A 2: GoTo E
mCmd = jj.Fmt_Str("""{0}"" a -p20071122 ""{1}"" ""{2}""", cPgm, mFfnZip, pFfn)
Shell mCmd, vbHide
'If Vba.Dir(mFfnZip) = "" Then ss.A 1, "Fail zipping (zip not found)", "pFfn", pFfn: Goto E
Exit Function
R: ss.R

E:
: ss.B cSub, cMod, "pFfn,pOvrWrt", pFfn, pOvrWrt
    Zip_By7Z = True
End Function
#If Tst Then
Function Zip_By7Zip_Tst() As Boolean
Dim mWb As Workbook: If jj.Crt_Wb(mWb, "c:\aa.xls") Then Stop
If jj.Cls_Wb(mWb, True) Then Stop
If Zip_By7Z("C:\aa.xls", True) Then Stop
End Function
#End If
Function UnZip_By7Z(pFfn$, Optional pOvrWrt As Boolean = False) As Boolean
'Aim: UnZip to {pFfn} from zip file in same directory
Const cSub$ = "UnZip_By7Z"
Const cPgm$ = "C:\Program Files\7-zip\7z.exe"
If Right(pFfn, 4) = ".zip" Then ss.A 1, "Cannot end with .zip", "pFfn", pFfn: GoTo E
Dim mCmd$, mFfnZip$
Dim mDir$, mFnn$, mExt$: If jj.Brk_Ffn_To3Seg(mDir$, mFnn, mExt, pFfn) Then ss.A 1: GoTo E
mFfnZip$ = jj.Cut_Ext(pFfn) & ".zip"
If VBA.Dir(mFfnZip) = "" Then ss.A 1, "No zip file exist", "mFfnZip", mFfnZip: GoTo E
If jj.Ovr_Wrt(pFfn, pOvrWrt) Then ss.A 2: GoTo E
mCmd = jj.Fmt_Str("""{0}"" x -p20071122 ""{1}"" ""{2}""", cPgm, mFfnZip, pFfn)
Shell mCmd, vbHide
'If Vba.Dir(pFfn) = "" Then ss.A 1, "Cannot extract pFfn$ from zip file", "pFfn", pFfn: Goto E
Exit Function
R: ss.R
E: UnZip_By7Z = True: ss.B cSub, cMod, "pFfn,pOvrWrt", pFfn, pOvrWrt
End Function
#If Tst Then
Function UnZip_By7Zip_Tst() As Boolean
If UnZip_By7Z("c:\aa.xls", True) Then Stop
End Function
#End If
