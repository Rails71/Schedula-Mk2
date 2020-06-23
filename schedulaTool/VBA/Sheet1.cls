VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub CommandButton1_Click()
    ' Parameters
    Dim season As String
    Dim fixtureFile As String
    Dim proxyAddr As String
    Dim useProxy As Boolean
    Dim username As String
    
    MsgBox ("Pull All")
    
    ' Get parameters
    season = Range("Season").Value
    fixtureFile = Range("FixtureFile").Value
    proxyAddr = Range("httpProxy").Value
    useProxy = Range("Enable_Proxy").Value
    username = Range("Username").Value
    
    ' Call the rotuine
    Call Module1.pullAll(fixtureFile, season, proxyAddr, useProxy, username)
End Sub

Private Sub CommandButton2_Click()
    ' Parameters
    Dim season As String
    Dim fixtureFile As String
    Dim proxyAddr As String
    Dim useProxy As Boolean
    Dim username As String
    Dim numberDays As String
    Dim startDay As String
    
    MsgBox ("Pull Some")
    
    ' Get parameters
    season = Range("Season").Value
    fixtureFile = Range("FixtureFile").Value
    proxyAddr = Range("httpProxy").Value
    useProxy = Range("Enable_Proxy").Value
    username = Range("Username").Value
    numberDays = Range("Days_to_pull").Value
    startDay = Format(Range("start_day").Value, "yyyy-mm-dd")
    
    ' Call the rotuine
    Call Module1.pullN(fixtureFile, season, proxyAddr, useProxy, _
                       username, numberDays, startDay)
End Sub

Private Sub CommandButton3_Click()
    ' Parameters
    Dim season As String
    Dim proxyAddr As String
    Dim useProxy As Boolean
    Dim username As String
    Dim peopleFile As String
    
    MsgBox ("Update People")
    
    ' Get parameters
    season = Range("Season").Value
    proxyAddr = Range("httpProxy").Value
    useProxy = Range("Enable_Proxy").Value
    username = Range("Username").Value
    peopleFile = Range("PeopleFile").Value
    
    ' Call the rotuine
    Call Module1.pullP(peopleFile, season, proxyAddr, useProxy, _
                       username)
End Sub

Private Sub CommandButton4_Click()
    Dim season As String
    Dim proxyAddr As String
    Dim useProxy As Boolean
    Dim username As String
    Dim peopleFile As String
    Dim fixtureFile As String
    Dim pushFile As String
    
    MsgBox ("Push")
    
    ' Get parameters
    season = Range("Season").Value
    fixtureFile = Range("FixtureFile").Value
    proxyAddr = Range("httpProxy").Value
    useProxy = Range("Enable_Proxy").Value
    username = Range("Username").Value
    peopleFile = Range("PeopleFile").Value
    pushFile = "push.csv"
    
    ' Export the appointments
    Call Module2.exportChanges
    
    ' Push the changes to schedula
    Call Module1.push(fixtureFile, peopleFile, pushFile, proxyAddr, useProxy, username)
End Sub

Private Sub CommandButton5_Click()
    Dim table As ListObject
    
    MsgBox ("Load to Sheet")
    
    ' Update connections
    ActiveWorkbook.Connections("Query - store").Refresh
    ActiveWorkbook.Connections("Query - people").Refresh
    
    ' Clear filters
    ActiveWorkbook.SlicerCaches("Slicer_Area").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Match_Status").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Organisation").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Competition_Name").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Competition_Group").ClearManualFilter
    
    ' Clear schedula table (except one row)
    Sheets("Schedula").Select
    Set table = ActiveSheet.ListObjects("schedula")
    With table.DataBodyRange
        If .Rows.count > 1 Then
            .Offset(1, 0).Resize(.Rows.count - 1, .Columns.count).Rows.Delete
        End If
    End With
    
    ' Clear fixtureid of row 1
    ' Range("A3").ClearContents
    ' Columns("A:A").ColumnWidth = 5
    ' Load in fixture ids
    Sheets("FixtureDB").Select
    ' Range("store[FixtureID]").Select
    ActiveSheet.ListObjects("store").ListColumns(1).DataBodyRange.Copy table.ListColumns(1).DataBodyRange
    
    Sheets("Schedula").Select

    ' Sort (date, venue, time)
    table.AutoFilter.ApplyFilter
    
End Sub
