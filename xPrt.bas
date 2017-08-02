Attribute VB_Name = "xPrt"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xPrt"
Function Prt_Str(pFno As Byte, pS) As Boolean
If pFno = 0 Then Debug.Print pS;: Exit Function
Print #pFno, pS;
End Function
Function Prt_Ln(pFno As Byte, Optional pS = "") As Boolean
If pFno = 0 Then Debug.Print pS: Exit Function
Print #pFno, pS
End Function
Function Prt_PDF(pFfnPDF$) As Boolean
Shell jj.Fmt_Str("""C:\Program Files\Adobe\Acrobat 7.0\Reader\AcroRd32.exe"" /p ""{0}""", pFfnPDF), vbMaximizedFocus
End Function

