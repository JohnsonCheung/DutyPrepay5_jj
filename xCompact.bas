Attribute VB_Name = "xCompact"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xCompact"
Function Compact_Db(pNmMdb$, Optional pKeepBackupLvl As Byte = 3) As Boolean
Const cSub$ = "Compact_CompactDb"
If Not jj.IsFfn(pNmMdb) Then ss.A 1: GoTo E
Dim mFfnn$: mFfnn = jj.Cut_Ext(pNmMdb)
On Error GoTo R
Dim mFfnCompact$: mFfnCompact = mFfnn & "_Compact.Mdb": If jj.Dlt_Fil(mFfnCompact) Then ss.A 2: GoTo E
g.gDbEng.CompactDatabase pNmMdb, mFfnn & "_Compact.Mdb"
If jj.Ren_ToBackup(pNmMdb, pKeepBackupLvl) Then ss.A 3: GoTo E
If jj.Ren_Fil(mFfnn & "_Compact.Mdb", pNmMdb) Then ss.A 4: GoTo E
Exit Function
R: ss.R
E: Compact_Db = True: ss.B cSub, cMod, "pNmMdb,pKeepBackupLvl", pNmMdb, pKeepBackupLvl
End Function
Function Db_Tst() As Boolean
Db_Tst = jj.Compact_Db("M:\07 ARCollection\ARCollection\WorkingDir\ARCollection_Data.mdb")
End Function
