Attribute VB_Name = "z"
Option Compare Text
Option Explicit
Const cMod$ = cLib & ".z"
'ARBal 4Var Fix SimRule
''4Var RPAAP, RPFAP(AP): amount payable (ie current outstanding).  RPAG(Gross) RPACR (Curr)
''Fix 193RecTblF0311Excl 20Rec0AG
'''20Rec0AG <>HKD, AG=0 ==> AG=ACR*CRR
''SimRule: ARBal_Cur ARBal_At
'''ARBal_Cur=@RPAN8, RPDCT, RPDOC, RPCRCD Sum(RPAAP & RPFAP)              WHERE RPAN8,RPAAP<>0 RPPST='A' RPDCTM Is Null AND RPDOCM=0
'''ARBal_At =@RPAN8, RPDCT, RPDOC, RPCRCD Sum(RPAAP & RPFAP=>OsBas OsCur) WHERE RPDGJ<={CurAsAtJdte} AND RPAN8<>0 (RPDCTM Not In ('RG','RQ') OR ISNULL(RPDCTM)) AND (RPSFXM<>'999' OR ISNULL(RPSFXM))
Function ZZZ() As Boolean
'If jj.Fnd_PrcBody_Tst Then Stop: GoTo E
'If jj.Brk_PrcBody_Tst Then Stop: GoTo E
If jj.Exp_Pgm_ToTbl_Tst Then Stop: GoTo E
'If jj.Run_Lgc_Tst Then Stop: GoTo E
'If jj.Run_Lgs_Tst Then Stop: GoTo E
'If jj.Run_Fb_Tst Then Stop: GoTo E
'If Rqp.Dta2Mdb_Tst Then Stop: GoTo E
'If jj.Run_Lgc_Tst Then Stop: GoTo E
'If jj.Bld_OdbcQs_ByAySelSql_Tst Then Stop: GoTo E
'If jj.Bld_LExpr_ByLpAp_Tst Then Stop: GoTo E
'If jj.Bld_OdbcQs_BySql_Tst Then Stop: GoTo E
'If jj.Bld_OdbcQs_Tst Then Stop: GoTo E
'If jj.Bld_Sql_Upd_ByRsUlSrc_Tst Then Stop: GoTo E
'If jj.Bld_Tbl_NRec2Lst_Tst Then Stop: GoTo E
'If jj.Cmp_TwoTbl_Tst Then Stop: GoTo E
'If jj.Compact_Db_Tst Then Stop: GoTo E
'If jj.Chk_Host_ByFrm_Tst Then Stop: GoTo E
'If jj.Run_Tp_Tst Then Stop: GoTo E
'If jj.Exp_SetNmtq2Dir_Tst Then Stop: GoTo E
'If jj.Cpy_Am_Tst Then Stop: GoTo E
'If jj.Cpy_Obj_ByLn_Tst Then Stop: GoTo E
'If jj.Cpy_Obj_ByPfx_Tst Then Stop: GoTo E
'If jj.Crt_Qry_ByDSN_Tst Then Stop: GoTo E
'If jj.Crt_PDF_FmWrd_Tst Then Stop: GoTo E
'If jj.Crt_Tbl_ByLnkLdb_Tst Then Stop: GoTo E
'If jj.Crt_Tbl_FmLnkMdb_Tst Then Stop: GoTo E
'If jj.Crt_Tbl_ForEdtTbl_Tst Then Stop: GoTo E
'If jj.Crt_TqRel_Tst Then Stop: GoTo E
'If jj.Crt_Xls_FmHost_ForEdt_Tst Then Stop: GoTo E
'If jj.Crt_Xls_FmNmt_ForEdt_Tst Then Stop: GoTo E
'If DlDta_Fm400BySql_Tst Then Stop: GoTo E
'If jj.Fmt_Tbl_Tst Then Stop: GoTo E
'If jj.GenXls_Tst Then Stop: GoTo E
'If jj.Fnd_Aim_Tst Then Stop: GoTo E
'If jj.Fnd_AnPrc_Tst Then Stop: GoTo E
'If jj.Fnd_Anq_ByNmqs_Tst Then Stop: GoTo E
'If jj.Fnd_CdMod_Tst Then Stop: GoTo E
'If jj.Fnd_Dte_Tst Then Stop: GoTo E
'If jj.Fnd_LoAyV_FmRs_Tst Then Stop: GoTo E
'If jj.Fnd_PrcBody_Tst Then Stop: GoTo E
'If jj.Fnd_ResStr_Tst Then Stop: GoTo E
'If ImpCus_Tst Then Stop: GoTo E
'If ImpXls_Tst Then Stop: GoTo E
'If IsEq_Tst Then Stop: GoTo E
'If RmvItm_InLst_Tst Then Stop: GoTo E
'If RunQry_ByAnq_Tst Then Stop: GoTo E
'If SetPdfPrt_Tst Then Stop: GoTo E
'If SndMail_Tst Then Stop: GoTo E
'If StrFormat_ByLn_Tst Then Stop: GoTo E
'If StrFormat_ByLp_Tst Then Stop: GoTo E
'If SubstWrd_Tst Then Stop: GoTo E
'If UlTbl_ToHost_Tst Then Stop: GoTo E
'If SetRgeVdt_ByLv_Tst Then Stop: GoTo E
'If JoinLst_Tst Then Stop: GoTo E
'If jj.Crt_Tbl_ByMgeNRec_To1Fld_Tst Then Stop: GoTo E
'If jj.Bld_LExpr_Tst Then Stop: GoTo E
Exit Function
E:
End Function
Function zGen_Doc() As Boolean
jj.Gen_Doc ' "Wrt*"
End Function

Function aa()
'If jj.Crt_Tbl_FmLoFld("#Tmp", "Id Long, AA Text, BB Text", 1) Then Stop
DoCmd.RunSQL "Insert into [#Tmp] (aa) values (1)"
Stop
E:
MsgBox Err.Description
Stop
End Function
