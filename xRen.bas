Attribute VB_Name = "xRen"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xRen"
Function Ren_Fil(pFfnFm$, pFfnTo$) As Boolean
Const cSub$ = "Ren_Fil"
If VBA.Dir(pFfnTo) <> "" Then ss.A 1, "To File Exist", , "pFfnFm,pFfnTo", pFfnFm, pFfnTo: GoTo E
On Error GoTo R
Name pFfnFm As pFfnTo
Exit Function
R: ss.R
E: Ren_Fil = True: ss.B cSub, cMod, "pFfnFm,pFfnTo", pFfnFm, pFfnTo
 End Function
Function Ren_Qry(pFmPfx$, pToPfx$) As Boolean
Dim iQry As QueryDef
Dim L%: L = Len(pFmPfx)
For Each iQry In CurrentDb.QueryDefs
    If Left(iQry.Name, L) = pFmPfx Then
        Debug.Print "Replacing Qry ... "; iQry.Name
        iQry.Name = pToPfx & mID$(iQry.Name, L + 1)
    End If
Next
End Function
Function Ren_QrySet(pQryPfx$, pBegNum As Byte, pEndNum As Byte, pToNum As Byte) As Boolean
If pToNum = pBegNum Then MsgBox "pToNum must <> pBegNum": Exit Function
If pEndNum < pBegNum Then MsgBox "pEndNum must > pBegNum": Exit Function
Dim J%
If pToNum > pBegNum Then
    For J = pEndNum To pBegNum Step -1
        jj.Ren_Qry pQryPfx & "_" & VBA.Format(J, "00"), pQryPfx & "_" & VBA.Format(J + pToNum - pBegNum, "00")
    Next
Else
    For J = pBegNum To pEndNum
        jj.Ren_Qry pQryPfx & "_" & VBA.Format(J, "00"), pQryPfx & "_" & VBA.Format(J + pToNum - pBegNum, "00")
    Next
End If
End Function
Function Ren_Tbl_ByNmt(pNmtFm$, pNmtTo$) As Boolean
Const cSub$ = "Ren_Tbl_ByNmt"
If jj.Dlt_Tbl(pNmtTo) Then ss.A 1, "pNmtTo cannot be deleted": GoTo E
On Error GoTo R
CurrentDb.TableDefs(pNmtFm).Name = pNmtTo
Exit Function
R: ss.R
E: Ren_Tbl_ByNmt = True: ss.B cSub, cMod, "pNmtFm,pNmtTo", pNmtFm, pNmtTo
End Function
Function Ren_Tbl_ByPfx(pFmPfx$, pToPfx$) As Boolean
Dim L%: L = Len(pFmPfx)
Dim iTbl As TableDef: For Each iTbl In CurrentDb.TableDefs
    If Left(iTbl.Name, L) = pFmPfx Then
        Debug.Print "Renaming ... "; iTbl.Name
        iTbl.Name = pToPfx & mID$(iTbl.Name, L + 1)
    End If
Next
End Function
Function Ren_ToBackup(pFb$, Optional pKeepBackupLvl As Byte = 3) As Boolean
Const cSub$ = "Ren_ToBackup"
If pKeepBackupLvl = 0 Then
    If jj.Dlt_Fil(pFb) Then ss.A 1: GoTo E
    Exit Function
End If
If pKeepBackupLvl > 9 Then pKeepBackupLvl = 9
Dim mFfnn$, mExt$: If jj.Brk_Ffn_To2Seg(mFfnn, mExt, pFb) Then ss.A 1: GoTo E
Dim mNxtFfnn$, mNxtBkNo As Byte: jj.Fnd_NxtBkFfnn mFfnn, mNxtFfnn, mNxtBkNo
If mNxtBkNo >= 10 Or mNxtBkNo >= pKeepBackupLvl Then
    If jj.Dlt_Fil(mNxtFfnn & mExt, True) Then ss.A 1: GoTo E
    If jj.Ren_Fil(pFb, mNxtFfnn & mExt) Then ss.A 2: GoTo E
    If jj.Set_FilRO(mNxtFfnn & mExt) Then ss.A 3: GoTo E
    Exit Function
End If
If VBA.Dir(mNxtFfnn & mExt) <> "" Then If jj.Ren_ToBackup(mNxtFfnn & mExt, pKeepBackupLvl) Then Exit Function
If jj.Ren_Fil(pFb, mNxtFfnn & mExt) Then ss.A 2: GoTo E
Exit Function
R: ss.R
E: Ren_ToBackup = True: ss.B cSub, cMod, "pFb,pKeepBackupLvl", pKeepBackupLvl
End Function
Function Ren_Tbl_ByLnk(pToPfx$) As Boolean
'Aim: Rename all linked table by adding {pToPfx}
Const cSub$ = "Ren_Tbl_ByLnk"
On Error GoTo R
If pToPfx = "" Then ss.A 1, "pToPfx cannot be blank": GoTo E
Dim iTbl As TableDef
For Each iTbl In CurrentDb.TableDefs
    If iTbl.Connect <> "" Then iTbl.Name = pToPfx & iTbl.Name
Next
Exit Function
R: ss.R
E: Ren_Tbl_ByLnk = True: ss.B cSub, cMod, "pToPfx", pToPfx
End Function

