Attribute VB_Name = "modResStr"
Option Compare Text
Option Explicit
Sub Shw_MnuWb()
'Sub Shw_MnuWb()
'Const cSub$ = "Shw_MnuWb"
'On Error GoTo R
'{0}.Show
'Exit Sub
'R: ss.R
'E: ss.B cSub, cMod
'End Sub
End Sub
Sub Shw_MnuWs()
'Sub Shw_MnuWs()
'Const cSub$ = "Shw_MnuWs"
'On Error GoTo R
'Dim mNmWs$: mNmWs = Excel.Application.ActiveSheet.CodeName
'Select Case mNmWs
'Case "Ws{N}": MnuWs{N}.Show
'End Select
'R: ss.R
'E: ss.B cSub, cMod
'End Sub
End Sub
Sub DtfTp()
'[DataTransferFromAS400]
'Version=2.0
'[HostInfo]
'HostFile={LIB}/IIC
'HostName={IP}
'[ClientInfo]
'ASCIITruncation=1
'ConvType={ConvTyp}
'CrtOpt=1
'FDFFile={FfnFDF}
'FDFFormat=0
'FileOps=33447039
'OutputDevice=2
'PCFile={FfnTar}
'PCFileType={PCFilTyp}
'SaveFDF={SavFDF}
'[SQL]
'EnableGroup=0
'GroupBy=
'Having=
'JoinBy=
'MissingFields=0
'OrderBy=
'SQLSelect={Sql}
'Select=
'Where=
'[Options]
'DateFmt=ISO
'DateSep=[/]
'DecimalSep=.
'IgnoreDecErr=1
'Lang=0
'LangID=
'SortSeq=0
'SortTable=
'TimeFmt=HMS
'TimeSep=[:]
'[HTML]
'AutoSize=0
'AutoSizeKB=128
'CapAlign=0
'CapIncNum=0
'CapSize=6
'CapStyle=1
'Caption=
'CellAlignN=0
'CellAlignT=0
'CellSize=6
'CellWrap=1
'Charset=big5
'ConvInd=0
'DateTimeLoc=0
'IncDateTime=0
'OverWrite=1
'RowAlignGenH=0
'RowAlignGenV=0
'RowAlignHdrH=0
'RowAlignHdrV=0
'RowStyleGen=1
'RowSytleHdr=1
'TabAlign=0
'TabBW=1
'TabCP=1
'TabCS=1
'TabCols=2
'TabMap=1
'TabRows=2
'TabWidth=100
'TabWidthP=0
'Template=
'TemplateTag=
'Title=
'UseTemplate=0
'[Properties]
'AutoClose=0
'AutoRun=0
'Check4Untrans=0
'Convert65535=0
'Notify=1
'SQLStmt=1
'ShowWarnings=0
'UseAlias=1
'UseCompression=1
'UserOption=0
'[LibraryList]
'Lib1=B {LIB}
End Sub
Private Sub GenDoc_FmtMod()
'''-- This is located in jj.mda  It is be used to put into the <Nam>_Doc.xls Sheet1 ---
'Dim x_MsAccess As Access.Application
'Public Property Get mMsAccess As Access.Application
'If TypeName(x_MsAccess) = "Nothing" Then
'    Set x_MsAccess = New Access.Application
'    Call x_MsAccess.OpenCurrentDatabase(Me.Cells(1, 1).Value)
'End If
'Set mMsAccess = x_MsAccess
'End Property
'Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
'Dim mNmm$, mLineNo&, J&
'For J = Range(Target.SubAddress).Row To 1 Step -1
'    If Not IsEmpty(Cells(J, 1).Value) Then
'        mNmm = Cells(J, 1).Value
'        Exit For
'    End If
'Next
'Dim mTyp As AcObjectType, mParam$
'mTyp = 5
'mParam = """" & mNmm & ""cComma & mTyp & cComma & Range(Target.SubAddress).Value
'Call mMsAccess.Eval("Line(" & mParam & ")") ' jj.Goto.Line
'End Sub
'
End Sub
Private Sub zzGenDoc_FmtQry()
'''-- This is located in jj.  It is be used to put into the <Nam>_Doc.xls!<Queires> Sheet2 ---
'Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
'Dim mQn$, J&, mRow
'Select Case Range(Target.SubAddress).Column
'Case 4, 5, 7
'Case Else
'   Exit Sub
'End Select
'
''Find mQn XXXX_00_0_XX
'
''0_XX
'mRow = Range(Target.SubAddress).Row
'If Range("D" & mRow).Value = "" Then
'    mQn = "0_" & Range("C" & mRow).Value
'Else
'    mQn = Range("D" & mRow).Value & "_" & Range("F" & mRow).Value
'End If
''Find 00 of 00_0_XX, which is <MajNo> at col B
'Dim mFound As Boolean: mFound = False
'For J = mRow To 1 Step -1
'    If Not IsEmpty(Cells(J, 2).Value) Then     'Column B look for <Maj#>
'        mQn = Format(Cells(J, 2).Value, "00") & "_" & mQn
'        mFound = True
'        Exit For
'    End If
'Next
'If Not mFound Then Stop
'
''Find XXXX of XXXX_00_0_XX, which <Nmqs> at col A
'mFound = False
'For J = mRow To 1 Step -1
'    If Not IsEmpty(Cells(J, 1).Value) Then     'Column A look for <Nmqs>
'        mQn = Cells(J, 1).Value & "_" & mQn
'        mFound = True
'        Exit For
'    End If
'Next
'If Not mFound Then Stop
'
'Select Case Range(Target.SubAddress).Column
'Case 4 ' Column D: <MinNo> -- View Definition
'    Call Sheet1.mMsAccess.Eval("GotoQryDef(""" & mQn & """)")
'Case 5 ' Column E -- View Data
'    Call Sheet1.mMsAccess.Eval("GotoQryView(""" & mQn & """)")
'    Exit Sub
'Case 7 ' Column G -- Update Remark
'    Dim mRmk$
'    mRmk = Range("H" & mRow).Value
'    Call Sheet1.mMsAccess.Eval("GotoQryRmk(""" & mQn & ""cComma"" & mRmk & """)")
'End Select
'End Sub
End Sub
