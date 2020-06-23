Attribute VB_Name = "Module1"
' Module 1
' Callbacks for user interaction
Option Explicit

' Calls the push command and updates the database
Public Sub push(fixtureFile As String, peopleFile As String, pushFile As String, _
                proxyAddress As String, useProxy As Boolean, _
                username As String)
                
    Dim command As String
    
    MsgBox ("fixFile: " & fixtureFile & vbCrLf & _
            "peopleFile: " & peopleFile & vbCrLf & _
            "pushFile: " & pushFile & vbCrLf & _
            "proxyAddr: " & proxyAddress & vbCrLf & _
            "useProxy: " & CStr(useProxy))
            
    command = """push"""
    command = addParameter(command, "-f")
    command = addParameter(command, (ActiveWorkbook.path & "\data\" & fixtureFile))
    command = addParameter(command, "-o")
    command = addParameter(command, (ActiveWorkbook.path & "\data\" & peopleFile))
    command = addParameter(command, "-i")
    command = addParameter(command, (ActiveWorkbook.path & "\data\" & pushFile))
    
    If useProxy Then
        command = addParameter(command, "-x")
        command = addParameter(command, proxyAddress)
    End If
    If username <> "" Then
        command = addParameter(command, "-u")
        command = addParameter(command, username)
    End If
    
    callSchedulaTool (command)
    
End Sub


' Calls the pullAll command and updates the database
Public Sub pullAll(fixtureFile As String, season As String, _
                   proxyAddress As String, useProxy As Boolean, _
                   username As String)
    Dim command As String
    
    MsgBox ("fixFile: " & fixtureFile & vbCrLf & _
            "season: " & season & vbCrLf & _
            "proxyAddr: " & proxyAddress & vbCrLf & _
            "useProxy: " & CStr(useProxy))
    
    ' Construct arguments
    command = """pullAll"""
    command = addParameter(command, "-f")
    command = addParameter(command, (ActiveWorkbook.path & "\data\" & fixtureFile))
    
    If season <> "" Then
        command = addParameter(command, "-s")
        command = addParameter(command, season)
    End If
    If useProxy Then
        command = addParameter(command, "-x")
        command = addParameter(command, proxyAddress)
    End If
    If username <> "" Then
        command = addParameter(command, "-u")
        command = addParameter(command, username)
    End If
    
    callSchedulaTool (command)
    
    ' TODO update from csv
End Sub

' Calls the pullN command and updates the database
Public Sub pullN(fixtureFile As String, season As String, _
                   proxyAddress As String, useProxy As Boolean, _
                   username As String, numberDays As String, startDay As String)
    Dim command As String
    
    MsgBox ("fixFile: " & fixtureFile & vbCrLf & _
            "season: " & season & vbCrLf & _
            "proxyAddr: " & proxyAddress & vbCrLf & _
            "useProxy: " & CStr(useProxy) & vbCrLf & _
            "numberDays: " & numberDays & vbCrLf & _
            "startDay: " & startDay)
    
    ' Construct arguments
    command = """pullN"""
    command = addParameter(command, "-f")
    command = addParameter(command, (ActiveWorkbook.path & "\data\" & fixtureFile))
    
    If season <> "" Then
        command = addParameter(command, "-s")
        command = addParameter(command, season)
    End If
    If useProxy Then
        command = addParameter(command, "-x")
        command = addParameter(command, proxyAddress)
    End If
    If username <> "" Then
        command = addParameter(command, "-u")
        command = addParameter(command, username)
    End If
    If numberDays <> "" Then
        command = addParameter(command, "-n")
        command = addParameter(command, numberDays)
    End If
    If startDay <> "" Then
        command = addParameter(command, "-N")
        command = addParameter(command, startDay)
    End If
    
    callSchedulaTool (command)
    
    ' TODO: update from csv
End Sub

Public Sub pullP(peopleFile As String, season As String, proxyAddr As String, useProxy As Boolean, username As String)
    Dim command As String
    
    MsgBox ("peopleFile: " & peopleFile & vbCrLf & _
            "season: " & season & vbCrLf & _
            "proxyAddr: " & proxyAddr & vbCrLf & _
            "useProxy: " & CStr(useProxy))
    
    ' Construct arguments
    command = """pullP"""
    command = addParameter(command, "-o")
    command = addParameter(command, (ActiveWorkbook.path & "\data\" & peopleFile))
    
    If season <> "" Then
        command = addParameter(command, "-s")
        command = addParameter(command, season)
    End If
    If useProxy Then
        command = addParameter(command, "-x")
        command = addParameter(command, proxyAddr)
    End If
    If username <> "" Then
        command = addParameter(command, "-u")
        command = addParameter(command, username)
    End If
    
    callSchedulaTool (command)
    
    ' TODO: update from csv
End Sub


' Calls the correct shell depending on operating system
Private Sub callSchedulaTool(command As String)
    Dim programName As String
    Dim pid As Variant
    
    ' Windows
    ' The tool is in .\schedulaTool\schedulaMain\schedulaMain.exe
    
    ' Mac
    ' TBA
    
    ' programName = ActiveWorkbook.path & "\schedulaTool\test\test.exe"
    programName = ActiveWorkbook.path & "\schedulaTool\schedulaMain\schedulaMain.exe"
    
    MsgBox programName & " " & command
    pid = Shell("""" & programName & """ " & command, vbNormalFocus)
    
    ' wait for the shell to complete
    
End Sub

Function addParameter(command As String, parameter As String)
    addParameter = command & " """ & parameter & """"
End Function
