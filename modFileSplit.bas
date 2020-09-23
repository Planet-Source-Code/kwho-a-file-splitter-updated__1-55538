Attribute VB_Name = "modFileSplit"
Option Explicit
Public Sub SplitFile(FileName As String, FileTitle As String, _
            SplitLength As Single, OutputName As String, _
             BinaryYN As Boolean, ForSpeed As Boolean, _
            pgbStatus As ProgressBar, lblStatus As Label, Optional DoneYN As Boolean)
    On Error Resume Next
    Dim Temp As String
    Dim TempB As Byte
    Dim TempA
    Dim FileData As String
    Dim EXtraSpace(1 To 10000)
    Dim SpaceCount As Single
    Dim I As Single
    Dim I2 As Single
    Dim Runs As Single
    Dim OutOf As Single
    Dim OutOfB As Single
    Dim OutOf2 As Single
    Dim PiecesDone As Single
    Dim FileLength As Single
    Dim SA As New FileSystemObject
    Dim BFile As File
    Dim DeleteYN As Boolean
    For I = 1 To 10000
        EXtraSpace(I) = ""
    Next I
    SpaceCount = 1
    I = 1
    
    lblStatus.Caption = "Grabbing file length..."
    DoEvents
    If BinaryYN = True Then
        Set BFile = SA.GetFile(FileName)
        I = CSng(BFile.Size)
        Set BFile = Nothing
    Else
        Open FileName For Input As #1
            Do Until EOF(1)
                Line Input #1, Temp
                I = I + 1
            Loop
        Close #1
    End If
    pgbStatus.Max = I
    OutOfB = I
    OutOf2 = I
    OutOf = Int(I / SplitLength)
    PiecesDone = 0
    I = 1
    I2 = 1
    If BinaryYN = True Then
        Open FileName For Binary Access Read As #1
            Do Until EOF(1)
                If I2 >= SplitLength Then
                    I2 = 1
                    PiecesDone = PiecesDone + 1
                End If
                If I Mod 100 = 0 Then
                    If ForSpeed = False Then
                        pgbStatus.Value = I
                        lblStatus.Caption = "Loading... " & PiecesDone + 1 & "/" & OutOf + 1 & " pieces (" & Int((I / OutOfB) * 100) & "%)"
                        DoEvents
                    End If
                End If
                If Len(FileData) >= 25000 Then
                    EXtraSpace(SpaceCount) = FileData
                    SpaceCount = SpaceCount + 1
                    FileData = ""
                End If
                Get #1, , TempB
                FileData = FileData & Chr(TempB)
                I = I + 1
                I2 = I2 + 1
            Loop
        Close #1
    Else
        Open FileName For Input As #1
            Do Until EOF(1)
                If I Mod 100 = 0 Then
                    If ForSpeed = False Then
                        pgbStatus.Value = I
                        lblStatus.Caption = "Loading line(s) " & I & "\" & OutOf2 & _
                            " (" & Int((I / OutOf2) * 100) & "%)"
                        DoEvents
                    End If
                End If
                If Len(FileData) >= 25000 Then
                    EXtraSpace(SpaceCount) = FileData
                    SpaceCount = SpaceCount + 1
                    FileData = ""
                End If
                Line Input #1, Temp
                FileData = FileData & Temp & vbCrLf
                I = I + 1
                
            Loop
        Close #1
    End If
    If SpaceCount > 1 Then
        pgbStatus.Max = SpaceCount + 1
        pgbStatus.Value = 1
        TempA = FileData
        FileData = ""
        For I = 1 To 10000
            pgbStatus.Value = I
            lblStatus.Caption = "Re-loading: " & I
            If ForSpeed = False Then DoEvents
            FileData = FileData & EXtraSpace(I)
            If ForSpeed = False Then DoEvents
            If I = SpaceCount + 1 Then Exit For
        Next I
        FileData = FileData & TempA
    End If
    FileLength = Len(FileData)
    Runs = Int(FileLength / SplitLength)
    pgbStatus.Max = Runs
    For I = 0 To Runs
        pgbStatus.Value = I
        If I = Runs Then
            If BinaryYN = False Then Open App.Path & "\" & OutputName & "." & I For Output As #1
            If BinaryYN Then Open App.Path & "\" & OutputName & "." & I For Binary Access Write As #1
                lblStatus.Caption = "Spliting && saving section: " & I + 1
                If ForSpeed = False Then DoEvents
                Temp = Mid$(FileData, I * SplitLength + 1, FileLength - (I * SplitLength))
                If ForSpeed = False Then DoEvents
                If BinaryYN = False Then Print #1, Temp
                If BinaryYN Then Put #1, , Temp
                SA.CopyFile App.Path & "\data1.dat", App.Path & "\" & OutputName & ".include.this"
                If ForSpeed = False Then DoEvents
            Close #1
        ElseIf I = 0 Then
            If BinaryYN = False Then Open App.Path & "\" & OutputName & "." & I For Output As #1
            If BinaryYN Then Open App.Path & "\" & OutputName & "." & I For Binary Access Write As #1
                lblStatus.Caption = "Spliting && saving section: " & I + 1
                If ForSpeed = False Then DoEvents
                Temp = Mid$(FileData, 1, SplitLength)
                If ForSpeed = False Then DoEvents
                If BinaryYN = False Then Print #1, Temp
                If BinaryYN Then Put #1, , Temp
                If ForSpeed = False Then DoEvents
            Close #1
        Else
            If BinaryYN = False Then Open App.Path & "\" & OutputName & "." & I For Output As #1
            If BinaryYN Then Open App.Path & "\" & OutputName & "." & I For Binary Access Write As #1
                lblStatus.Caption = "Spliting && saving section: " & I + 1
                If ForSpeed = False Then DoEvents
                Temp = Mid$(FileData, I * SplitLength + 1, SplitLength)
                If ForSpeed = False Then DoEvents
                If BinaryYN = False Then Print #1, Temp
                If BinaryYN Then Put #1, , Temp
                If ForSpeed = False Then DoEvents
            Close #1
        End If
    Next I
    Open App.Path & "\" & OutputName & "-restore.bat" For Output As #1
        Print #1, MakeBatFile(CInt(Runs), OutputName, FileTitle)
    Close #1
    DoneYN = True
End Sub


Public Function MakeBatFile(Runs As Integer, OutputName As String, FileTitle As String) As String
    Dim Temp As String
    Dim DeleteYN As Boolean
    Dim I As Integer
    Dim FirstStep As Boolean
    Temp = "echo off" & vbCrLf & _
        "cls" & vbCrLf & _
        "echo Restore - by kg_prog..." & vbCrLf & _
        "ren " & OutputName & ".include.this " & OutputName & ".include.this.exe" & vbCrLf & _
        "if exist ""%1" & FileTitle & """ del ""%1" & FileTitle & """" & vbCrLf & _
        "echo Restoring " & FileTitle & "..." & vbCrLf & _
        "call " & OutputName & ".include.this.exe" & vbCrLf & _
        "copy include.this ""%1" & FileTitle & """" & vbCrLf & _
        "if exist .delete goto deleteIt" & vbCrLf
    '====
DoIt:
    If DeleteYN = True Then
        Temp = Temp & ":deleteIt" & vbCrLf
    End If
    If DeleteYN = True Then Temp = Temp & "del include.this " & vbCrLf
    For I = 0 To Runs
        Temp = Temp & "copy /b /v ""%1" & FileTitle & """ + """ & OutputName & "." & I & """ ""%1" & FileTitle & """" & vbCrLf
        If DeleteYN = True Then
            Temp = Temp & "del " & OutputName & "." & I & vbCrLf & "echo " & OutputName & "." & I & " deleted." & vbCrLf
        End If
    Next I
    Temp = Temp & "echo Files copied." & vbCrLf
    If DeleteYN Then
        Temp = Temp & "del " & OutputName & ".include.this.exe" & vbCrLf & "del .delete" & vbCrLf & "del include.this" & vbCrLf & "del " & OutputName & "-restore.bat" & vbCrLf
    End If
    If FirstStep = False Then
        Temp = Temp & "goto endOfFile" & vbCrLf
        FirstStep = True
        DeleteYN = True
        GoTo DoIt
    End If
    Temp = Temp & ":endOfFile" & vbCrLf & "if exist .delete del .delete" & vbCrLf & "del include.this" & vbCrLf & _
            "ren " & OutputName & ".include.this.exe " & OutputName & ".include.this"
    MakeBatFile = Temp
End Function
