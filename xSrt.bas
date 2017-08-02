Attribute VB_Name = "xSrt"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xSrt"
Function Srt_Ay(pAy$(), oAy$()) As Boolean
Const cSub$ = "Srt_Ay"
Const cNmtTmp$ = "TmpSrt"
If jj.Siz_Ay(pAy) = 0 Then oAy = pAy: Exit Function
On Error GoTo R
If jj.Dlt_Tbl(cNmtTmp) Then ss.A 1: GoTo E
If jj.Run_Sql(jj.Fmt_Str("Create table {0} (aa Text(255))", cNmtTmp)) Then ss.A 2: GoTo E
With CurrentDb.OpenRecordset("Select * from " & cNmtTmp)
    Dim J%: For J = LBound(pAy) To UBound(pAy)
        .AddNew
        !aa.Value = pAy(J)
        .Update
    Next
    .Close
End With
ReDim oAy(LBound(pAy) To UBound(pAy))
With CurrentDb.OpenRecordset("Select * from " & cNmtTmp & " order by aa")
    J = LBound(pAy)
    While Not .EOF
        oAy(J) = !aa
        J = J + 1
        .MoveNext
    Wend
    .Close
End With
If jj.Dlt_Tbl(cNmtTmp) Then ss.A 3: GoTo E
Exit Function
R: ss.R
E: Srt_Ay = True: ss.B cSub, cMod
End Function
Function Srt_Ay_Tst() As Boolean
Dim mA$(2), mB$()
mA(0) = 10
mA(1) = 9
mA(2) = 8
If jj.Srt_Ay(mA, mB) Then Stop
Debug.Print mB(0)
Debug.Print mB(1)
Debug.Print mB(2)
Debug.Print jj.Siz_Ay(mB)
End Function
Function Srt_Coll(oColl As VBA.Collection, pColl As VBA.Collection) As Boolean
Const cSub$ = "Srt_SortColl"
Const cNmtTmp$ = "#Tmp"
If jj.IsAcs Then
    If jj.Dlt_Tbl(cNmtTmp) Then ss.A 1: GoTo E
    If jj.Run_Sql(jj.Fmt_Str("Create table {0} (aa Text(255))", cNmtTmp)) Then ss.A 1: GoTo E
    Dim iStr: For Each iStr In pColl
        If jj.Run_Sql("insert into [" & cNmtTmp & "] (aa) values('" & iStr & "')") Then ss.A 2: GoTo E
    Next
    Set oColl = New VBA.Collection
    With CurrentDb.OpenRecordset("Select * from [" & cNmtTmp & "] order by aa")
        While Not .EOF
            oColl.Add CStr(!aa.Value)
            .MoveNext
        Wend
        .Close
    End With
    jj.Dlt_Tbl cNmtTmp
    Exit Function
End If
#If Xls Then
If jj.IsXls Then
    Sheet1.Columns(1).Delete
    Dim iRno&: iRno = 0
    For Each iStr In pColl
        iRno = iRno + 1
        Sheet1.Cells(iRno, 1).Value = iStr
    Next
    Sheet1.Range("A1:A" & iRno).Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Set oColl = New VBA.Collection
    For iRno = 1 To iRno
        oColl.Add Sheet1.Cells(iRno, 1).Value
    Next
    Exit Function
End If
#End If
Exit Function
R: ss.R
E: Srt_Coll = True: ss.B cSub, cMod, "pColl", jj.ToStr_Coll(pColl)
End Function
Function Srt_Lv$(pLv$)
Const cSub$ = "Srt_Lv"
On Error GoTo R
If pLv$ = "" Then Srt_Lv = "": Exit Function
Dim mAy1$(): mAy1 = Split(pLv, cComma)
Dim mAy2$(): If jj.Srt_Ay(mAy1, mAy2) Then ss.A 1, "Cannot sort array": GoTo E
Srt_Lv = Join(mAy2, cComma)
Exit Function
R: ss.R
E: ss.B cSub, cMod, "pLv", pLv
Srt_Lv = ""
End Function
Function Srt_Pt(pPt As PivotTable) As Boolean
Dim iPf As PivotField
For Each iPf In pPt.PivotFields
    With iPf
        If .Name <> "Data" Then .AutoSort xlAscending, .Name
    End With
Next
End Function


