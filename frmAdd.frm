VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAdd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add file"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3825
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   3825
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdSplit 
      Caption         =   "Add File To List"
      Default         =   -1  'True
      Height          =   855
      Left            =   2040
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "File to split"
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Split file into"
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   600
      Width           =   3855
      Begin VB.TextBox txtFileSize 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "1440 (default)"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtPiece 
         Height          =   285
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "KB, or by pieces:"
         Height          =   195
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Temporary File Name"
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   1200
      Width           =   1935
      Begin VB.TextBox txtTempName 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Text            =   "file"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Optimized for speed"
      Height          =   615
      Left            =   2160
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CheckBox chkSpeed 
         Caption         =   "Yes"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   1320
   End
   Begin MSComDlg.CommonDialog comFile 
      Left            =   1560
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "File To Split"
      Filter          =   "(*.*)|*.*|Text (*.txt)|*.txt"
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Running As Boolean
Dim TimeAmount As Integer
Public IndexNum As Integer
Private Sub chkSpeed_Click()
    If chkSpeed.Value = 1 Then
        MsgBox "Note: this may cause your computer to freeze."
    End If
End Sub

Private Sub cmdBrowse_Click()
    comFile.ShowOpen
    If comFile.FileTitle = "" Then Exit Sub
    txtFile.Text = comFile.FileName
End Sub
Private Sub cmdSplit_Click()
    Dim SplitSize As Single
    Dim SA As New FileSystemObject
    Dim BFile As File
    Dim Temp() As String
    Dim TempItem As ListItem
    Dim BinaryYorN As Boolean
    If txtFile.Text = "" Then GoTo NotComplete
    If txtFileSize.Text = "" And txtPiece.Text = "" Then GoTo NotComplete
    If txtTempName.Text = "" Then GoTo NotComplete
    If txtFileSize.Text <> "" And txtPiece.Text <> "" Then GoTo NotComplete
    
    Temp = Split(txtTempName.Text, " ")
    If UBound(Temp) > 0 Then
        MsgBox "The temp name cannot have a ' ' (space) in it."
        Exit Sub
    End If
    If txtFileSize = "1440 (default)" Then txtFileSize = "1440"
    If txtFileSize.Text <> "" Then
        SplitSize = txtFileSize.Text
        SplitSize = SplitSize * 1024
    Else
        Set BFile = SA.GetFile(txtFile.Text)
        SplitSize = CSng(BFile.Size)
        Set BFile = Nothing
        SplitSize = SplitSize \ CSng(txtPiece.Text) + 1
    End If
    If LCase(Right$(txtFile.Text, 3)) <> "txt" Then BinaryYorN = True
    Set TempItem = frmMulti.lsvQue.ListItems(IndexNum)
    TempItem.Text = comFile.FileTitle
    TempItem.SubItems(1) = BinaryYorN
    TempItem.SubItems(2) = SplitSize
    TempItem.SubItems(3) = txtTempName.Text
    TempItem.SubItems(4) = txtFile.Text
    Set TempItem = Nothing
        
    Unload Me
    Exit Sub
NotComplete:
    MsgBox "Not every field is filled in\too many fields filled in..."
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Running = True Then End
End Sub


Private Sub tmrTime_Timer()
    TimeAmount = TimeAmount + 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
