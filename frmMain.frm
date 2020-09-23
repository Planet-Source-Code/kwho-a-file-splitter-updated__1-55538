VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Splitter - kg_prog"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   3240
   End
   Begin VB.Frame Frame6 
      Caption         =   "Optimized for speed"
      Height          =   615
      Left            =   1920
      TabIndex        =   17
      Top             =   1920
      Width           =   1935
      Begin VB.CheckBox chkSpeed 
         Caption         =   "Yes"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Temporary File Name"
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   1920
      Width           =   1935
      Begin VB.TextBox txtTempName 
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Text            =   "file"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Split file into"
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   1320
      Width           =   3855
      Begin VB.TextBox txtPiece 
         Height          =   285
         Left            =   2640
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtFileSize 
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "KB, or by pieces:"
         Height          =   195
         Left            =   1320
         TabIndex        =   14
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "File to split"
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   3855
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Delete Split Files After Use"
      Height          =   615
      Left            =   1680
      TabIndex        =   6
      Top             =   720
      Width           =   2175
      Begin VB.OptionButton optYes 
         Caption         =   "Yes"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optNo 
         Caption         =   "No"
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Type:"
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   1695
      Begin VB.OptionButton optBinary 
         Caption         =   "Binary"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optText 
         Caption         =   "Text"
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog comFile 
      Left            =   1320
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "File To Split"
      Filter          =   "(*.*)|*.*|Text (*.txt)|*.txt"
   End
   Begin MSComctlLib.ProgressBar pgbStatus 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdSplit 
      Caption         =   "Split File"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Idle..."
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Running As Boolean
Dim TimeAmount As Integer
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
    If txtFile.Text = "" Then GoTo NotComplete
    If txtFileSize.Text = "" And txtPiece.Text = "" Then GoTo NotComplete
    If txtTempName.Text = "" Then GoTo NotComplete
    If txtFileSize.Text <> "" And txtPiece.Text <> "" Then GoTo NotComplete
    
    Temp = Split(txtTempName.Text, " ")
    If UBound(Temp) > 0 Then
        MsgBox "The temp name cannot have a ' ' (space) in it."
        Exit Sub
    End If
    
    If txtFileSize.Text <> "" Then
        SplitSize = txtFileSize.Text
        SplitSize = SplitSize * 1024
    Else
        Set BFile = SA.GetFile(txtFile.Text)
        SplitSize = CSng(BFile.Size)
        Set BFile = Nothing
        SplitSize = SplitSize \ CSng(txtPiece.Text) + 1
    End If
    
    cmdSplit.Enabled = False
    tmrTime.Enabled = True
    Running = True
    SplitFile comFile.FileName, comFile.FileTitle, _
        SplitSize, txtTempName.Text, optYes, optBinary, _
        chkSpeed, pgbStatus, lblStatus
    Running = False
    tmrTime.Enabled = False
    MsgBox "Time to split: " & Int(TimeAmount / 60) & " min(s) & " & (TimeAmount Mod 60) & " sec(s)"
    cmdSplit.Enabled = True
        
    Unload Me
    Me.Show
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
