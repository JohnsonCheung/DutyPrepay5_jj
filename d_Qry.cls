VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "d_Qry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit
Const cMod$ = cLib & ".d_Qry"
Public NmQs$, Maj As Byte, Min As Byte, Rest$, Typ As DAO.QueryDefTypeEnum, Sql$, Des$, LnTbl$
Function Nmq$()
Nmq = NmQs & "_" & Format(Maj, "00") & "_" & Min & "_" & Rest
End Function
Function Brk_Nmqs(pNmq$) As Boolean
Const cSub$ = "Brk_Nmqs"
'Aim: Break {pNmq} in 4 segments of XXXX_NN_N_xxxx, return XXXX else return ""
Dim mMaj$, mMin$: If jj.Brk_Str_To4Seg(NmQs, mMaj, mMin, Rest, pNmq, "_") Then ss.A 1: GoTo E
If Len(mMaj) <> 2 Then ss.A 2, "Maj should be 2 chr": GoTo E
Maj = Val(mMaj): If VBA.Format(Maj, "00") <> mMaj Then ss.A 1, "Maj should be 2 digits: GoTo E"
If Len(mMin) <> 1 Then ss.A 1, "Min should be 1 chr", "pNmq", pNmq: GoTo E
Min = Val(mMin): If VBA.Format(Min, "0") <> mMin Then ss.A 2, "Min should be 1 digits": GoTo E
Exit Function
E:
: ss.B cSub, cMod, "pNmq", pNmq
    Brk_Nmqs = True
End Function
Function CrtTbl(Optional pNmt$ = "#OldQry") As Boolean
Const cSub$ = "CrtTbl"
If jj.Crt_Tbl_FmLoFld(pNmt, "Fb Text 255,NmQs Text 30,Maj Byte, Min Byte, Rest Text 50,NmTypQry Text 15,Sql Memo, LnTbl Memo, DesOldQry Memo") Then ss.A 1: GoTo E
Exit Function
R: ss.R

E:
: ss.B cSub, cMod, "pNmt", pNmt
    CrtTbl = True
End Function
Function InsAy(pAyDQry() As d_Qry, pFb$, pNmtDQry$) As Boolean
'Aim: Insert pAyDQry to pNmtDQry
Const cSub$ = "InsAy"
Dim J%, mN%: mN = jj.Siz_AyDQry(pAyDQry)
If mN = 0 Then Exit Function
On Error GoTo R
Dim mRs As DAO.Recordset: Set mRs = CurrentDb.TableDefs(pNmtDQry).OpenRecordset
For J = 0 To mN - 1
    If pAyDQry(J).Ins(mRs, pFb) Then ss.A 1: GoTo E
Next
GoTo X
R: ss.R

E:
: ss.B cSub, cMod, "pAyDQry(),pFb,pNmtDQry", "..", pFb, pNmtDQry
    InsAy = True
X:
    jj.Cls_Rs mRs
End Function
Function Ins(pRs As DAO.Recordset, pFb$) As Boolean
Const cSub$ = "Ins"
With pRs
    .AddNew
    !Fb = pFb
    !NmQs = NmQs
    !Maj = Maj
    !Min = Min
    !Rest = Rest
    !NmTypQry = jj.ToStr_TypQry(Typ)
    !Sql = Sql
    !LnTbl = LnTbl
    !DesOldQry = Des
    .Update
End With
End Function
Function WrtHdr(pF As Byte) As Boolean
Const cSub$ = "WrtHdr"
On Error GoTo R
Write #pF, "Fb", "NmQs", "Maj", "Min", "Rest", "NmTypQry", "Sql", "LnQDpd", "DesOldQry"
Exit Function
R: ss.R

E:
: ss.B cSub, cMod, "pF", pF
    WrtHdr = True
End Function
Function Wrt(pF As Byte, pFb$) As Boolean
Const cSub$ = "Wrt"
On Error GoTo R
Write #pF, pFb, NmQs, Maj, Min, Rest, jj.ToStr_TypQry(Typ), Sql, LnTbl
Exit Function
R: ss.R

E:
: ss.B cSub, cMod, "pF", pF
    Wrt = True
End Function

