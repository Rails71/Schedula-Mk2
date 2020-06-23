Attribute VB_Name = "Module2"
Option Explicit

Public Sub exportChanges()
'
' Export the changes in the schedula table

    Dim diff As Boolean
    Dim i As Variant
    Dim r As String
    Dim ar1 As String
    Dim ar2 As String
    Dim m As String
    Dim a As String
    Dim r4 As String
    Dim fixID As String
    Dim count As Long
    Dim path As String
    count = 1
    
    ' Clear output sheet
    Call clearOutput
    
    ' Goto the output sheet
    Sheets("Output").Select
    Range("A1").Select
    
    ' Write the header
    Cells(1, 1) = "FixtureID"
    Cells(1, 2) = "R"
    Cells(1, 3) = "AR1"
    Cells(1, 4) = "AR2"
    Cells(1, 5) = "M"
    Cells(1, 6) = "A"
    Cells(1, 7) = "4"
    
    ' Get updated fixtures
    For i = 1 To Range("schedula").Rows.count
        diff = Range("schedula[diff]")(i)
        
        If diff Then
            r = Range("schedula[Referee]")(i)
            ar1 = Range("schedula[AR1]")(i)
            ar2 = Range("schedula[AR2]")(i)
            m = Range("schedula[Mentor]")(i)
            a = Range("schedula[Assessor]")(i)
            r4 = Range("schedula[4th Official]")(i)
            fixID = Range("schedula[FixtureID]")(i)
            ' MsgBox ("Row " & CStr(i) & " is diff" & vbCrLf & _
                    "r " & r & vbCrLf & _
                    "ar1 " & ar1 & vbCrLf & _
                    "ar2 " & ar2 & vbCrLf & _
                    "m " & m & vbCrLf & _
                    "a " & a & vbCrLf & _
                    "r4 " & r4 & vbCrLf & _
                    "fID " & fixID)
                    
            ' write to output
            count = count + 1
            Cells(count, 1) = fixID
            Cells(count, 2) = r
            Cells(count, 3) = ar1
            Cells(count, 4) = ar2
            Cells(count, 5) = m
            Cells(count, 6) = a
            Cells(count, 7) = r4
        End If
    Next i

    ' Save as a csv
    path = ActiveWorkbook.path
    Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs Filename:=(path & "\data\push.csv"), FileFormat:=xlCSV, _
        CreateBackup:=False
    ActiveWindow.Close
    
    ' push output to schedula
    
End Sub

Sub createFixtureDB()
Attribute createFixtureDB.VB_ProcData.VB_Invoke_Func = " \n14"
'
' createFixtureDB Macro
'

'
    ActiveWorkbook.Queries.Add Name:="store", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Csv.Document(File.Contents(""C:\temp\schedulaMk2_git\schedulaTool\data\store.csv""),[Delimiter="","", Columns=22, Encoding=65001, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"",{{""FixtureID"", Int64.Type}, {" & _
        """OrgID"", Int64.Type}, {""OrgName"", type text}, {""SeasonID"", Int64.Type}, {""SeasonName"", Int64.Type}, {""WeekID"", type text}, {""WeekName"", type text}, {""Competition"", type text}, {""Date"", type date}, {""day"", type text}, {""Time"", type time}, {""Home"", type text}, {""Away"", type text}, {""Ground"", type text}, {""Referee"", type text}, {""AR1"", typ" & _
        "e text}, {""AR2"", type text}, {""Mentor"", type text}, {""Assessor"", type text}, {""4th Official"", type text}, {""Other"", type text}, {""Status"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    ActiveWorkbook.Worksheets.Add
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=store;Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [store]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "store"
        .Refresh BackgroundQuery:=False
    End With
    Sheets("FixtureDB").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "FixtureDB"
    Range("store[[#Headers],[FixtureID]]").Select
End Sub
Sub updateConnections()
Attribute updateConnections.VB_ProcData.VB_Invoke_Func = " \n14"
'
' updateConnections Macro
'

'
    ActiveWorkbook.Connections("Query - store").Refresh
    ActiveWorkbook.Connections("Query - people").Refresh
End Sub
Sub clearOutput()
Attribute clearOutput.VB_ProcData.VB_Invoke_Func = " \n14"
'
' clearOutput Macro
'

'
    Sheets("Output").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
End Sub

Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro5 Macro
'

'
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Workbooks.Add
    Windows("Schedula Mk2.xlsm").Activate
    Selection.Copy
    Windows("Book1").Activate
    ActiveSheet.Paste
    Application.Left = 1531
    Application.Top = 95.5
    Application.CutCopyMode = False
    ChDir "C:\temp"
    ActiveWorkbook.SaveAs Filename:="C:\temp\temp.csv", FileFormat:=xlCSVUTF8, _
        CreateBackup:=False
    ActiveWindow.Close
End Sub
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'

'
    Sheets("FixtureDB").Select
    ActiveWorkbook.RefreshAll
    Sheets("Schedula").Select
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWorkbook.SlicerCaches("Slicer_Area").ClearManualFilter
    ActiveWindow.SmallScroll Down:=-15
    ActiveWindow.SmallScroll ToRight:=-5
    ActiveCell.Offset(0, -38).Columns("A:A").EntireColumn.EntireColumn.AutoFit
    ActiveWindow.ScrollColumn = 1
    ActiveCell.Offset(-35, -38).Range("A1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    Selection.ListObject.ListRows(2).Delete
    Selection.ListObject.ListRows(2).Delete
    
    ActiveCell.Offset(-1, 0).Range("A1").Select
    Selection.ClearContents
    Sheets("FixtureDB").Select
    ActiveWindow.SmallScroll Down:=-33
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Schedula").Select
    ActiveSheet.Paste
    ActiveCell.Columns("A:A").EntireColumn.ColumnWidth = 0
    Range("O3").Select
    ActiveWindow.SmallScroll Down:=-30
End Sub
Sub Macro7()
Attribute Macro7.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro7 Macro
'

'
    ActiveWorkbook.SlicerCaches("Slicer_Area").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Organisation").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Competition_Name").ClearManualFilter
    Columns("A:A").ColumnWidth = 10.14
    Range("A4").Select
    Range("A4").Select
End Sub
Sub Macro8()
Attribute Macro8.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro8 Macro
'

'
    Columns("A:A").ColumnWidth = 5.57
    Columns("A:A").ColumnWidth = 0
End Sub
Sub Macro9()
Attribute Macro9.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro9 Macro
'

'
    Range("O3").Select
End Sub
Sub Macro10()
Attribute Macro10.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro10 Macro
'

'
    Application.CutCopyMode = False
    ActiveSheet.ListObjects("schedula").AutoFilter.ApplyFilter
    With ActiveWorkbook.Worksheets("Schedula").ListObjects("schedula").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Save
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    With ActiveWorkbook.SlicerCaches("Slicer_Area")
        .SlicerItems("Hills").Selected = True
        .SlicerItems("Not Assigned").Selected = False
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Organisation")
        .SlicerItems("NPLSA").Selected = True
        .SlicerItems("FFSA").Selected = False
    End With
    Range("O29").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    ActiveWorkbook.SlicerCaches("Slicer_Area").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Organisation").ClearManualFilter
    Range("O4").Select
    ActiveSheet.ListObjects("schedula").AutoFilter.ApplyFilter
    With ActiveWorkbook.Worksheets("Schedula").ListObjects("schedula").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Macro11()
Attribute Macro11.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro11 Macro
'

'
    ActiveWorkbook.SlicerCaches("Slicer_Area").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Match_Status").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Organisation").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Competition_Name").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Competition_Group").ClearManualFilter
    ActiveSheet.ListObjects("schedula").AutoFilter.ApplyFilter
    With ActiveWorkbook.Worksheets("Schedula").ListObjects("schedula").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Macro12()
Attribute Macro12.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro12 Macro
'

'
    Windows("push.csv").Activate
    ChDir "C:\temp\schedulaMk2_git\data"
    ActiveWorkbook.SaveAs Filename:="C:\temp\schedulaMk2_git\data\push.csv", _
        FileFormat:=xlCSV, CreateBackup:=False
    Windows("Schedula Mk2.xlsm").Activate
End Sub
