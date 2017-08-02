Attribute VB_Name = "Typ"
Option Compare Text
Option Explicit
'Public Type tDQry
'    Nmqs As String
'    Maj As Byte
'    Min As Byte
'    Rest As String
'    Typ As DAO.QueryDefTypeEnum
'    Sql As String
'    Rmk As String
'    LnQDpd As String
'End Type
Public Type tRgeCno
    Fm As Byte
    To As Byte
End Type
Public Type tGenWs
    Pt_SqlStr As String
    Pt_Nam As String
    PivotRows As String
    PivotColumns As String
    PivotData As String
    RowHeight As String
    ColWidth As String
    ColWidth_Default As Byte
    RowGrand As Boolean
    ColumnGrand As Boolean
    HideRows As String
    FreezeAt As String
    NmWsNew As String
End Type
'=====================================
Public Type tRgeRno
    Fm As Long
    To As Long
End Type
'=====================================
Public Type tSq
    c1 As Byte
    c2 As Byte
    r1 As Long
    r2 As Long
End Type
'=====================================
Public Type tFmtTbl_Def
    IsSepLin As Boolean '   Default is True: to turn off: After Top Right <Tbl>: NoSepLin
    Nmtq As String
    Sq As tSq           ' The square holding the definition
    RgeDta As Range
    SqDta As tSq
    NLvl As Byte
    Qt As QueryTable
    RgeCno_Lvl1 As tRgeCno
    RgeCno_Lvl2 As tRgeCno
    RgeCno_Lvl3 As tRgeCno
    RgeCno_Lvl4 As tRgeCno
    AyCnoSubTot() As Byte
    AyCnoAvg() As Byte
    AyCnoFormula() As Byte
    AyFormula() As String
    AyCnoVLine_Right() As Byte
    AyCnoVLine_Left() As Byte
    AyCnoVLine_RightMedium() As Byte
    AyCnoVLine_LeftMedium() As Byte
    AyRgeCno_Merge() As tRgeCno
    CnoCnt As Byte              ' Column with subtotal of count, Key Word = Cnt
End Type
'=====================================
Public Type tRpt
    RptId As Byte
    NmRptSht As String
    NmRpt As String
    FmtStr_FnTo As String
    QryPrm As String
    LnwsRmv As String
    HidePfLst_ThisNmSess As String
    HidePfLst_ThisSess As String
    HidePfLst_OtherSess As String
    NmDta As String
    EachNmFld As String
    EachLnwsRmv As String
    EachHidePfLst_ThisSess As String
    EachHidePfLst_OtherSess As String
    EachSql As String
End Type
'=====================================
Public Type tDocPrp
    NmRpt As String
    NmRptSht As String
    NmSess As String
    NmData As String
    ExtraPrm As String
End Type
'=====================================
Public Type tEnv
    To As String
    Fm As String
    CC As String
    Subj As String
    Body As String
    Ffn As String
End Type
'=====================================
Public Type tMap
    F1 As String
    F2 As String
End Type
'=====================================
Public Type tNm2V
    Nm As String
    NewV As Variant
    OldV As Variant
End Type
'=====================================
Public Type tYYYYMM
    YYYY As Integer
    MM As Byte
End Type
