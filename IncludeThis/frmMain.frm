VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "<<>>  Include.this.exe..."
   ClientHeight    =   285
   ClientLeft      =   2850
   ClientTop       =   3000
   ClientWidth     =   3240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   285
   ScaleWidth      =   3240
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim MsgResponse
    MsgResponse = MsgBox("Delete the pieces of the split file after reconnecting them?", vbYesNo, "File Splitter")
    If MsgResponse = vbYes Then
        Open App.Path & "\include.this" For Output As #1
        Close #1
        Open App.Path & "\.delete" For Output As #1
        Close #1
        End
    Else
        Open App.Path & "\include.this" For Output As #1
        Close #1
        End
    End If
End Sub
