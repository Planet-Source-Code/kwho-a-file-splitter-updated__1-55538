VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMulti 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi File-Splitting"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7140
   Icon            =   "frmMulti.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrOverAll 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   840
   End
   Begin VB.CheckBox chkRestoreAll 
      Caption         =   "Create a 'Restore All' program"
      Height          =   195
      Left            =   4680
      TabIndex        =   7
      Top             =   120
      Width           =   2415
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   240
   End
   Begin MSComctlLib.ProgressBar pgbStatus 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   4440
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Splitting!"
      Default         =   -1  'True
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   3840
      Width           =   3735
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Edit Selected File"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Choose file to add"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin MSComctlLib.ListView lsvQue 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Title"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Binary"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Split size"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Temp Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "File Name"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComDlg.CommonDialog comFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "File To Split"
      Filter          =   "(*.*)|*.*|Text (*.txt)|*.txt"
   End
   Begin VB.Label lblWorkingOn 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   7200
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label lblStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Line Line2 
      X1              =   3240
      X2              =   3240
      Y1              =   3120
      Y2              =   4440
   End
   Begin VB.Line Line1 
      X1              =   3240
      X2              =   7080
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Que Line:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   690
   End
End
Attribute VB_Name = "frmMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Running As Boolean
Dim TimeAmount As Integer
Dim OverAllTime As Integer
Private Sub cmdAdd_Click()
    Dim TempS() As String
    Dim Temp As String
    Dim TempI As Double
    Dim TempItem As ListItem
    Set TempItem = lsvQue.ListItems.Add
    comFile.ShowOpen
    If comFile.FileName = "" Then Exit Sub
    TempS = Split(comFile.FileName, "\")
    Temp = Left$(TempS(UBound(TempS)), Len(TempS(UBound(TempS))) - 4)
    TempItem.SubItems(3) = Temp & "TEMP"
    If LCase(Right$(comFile.FileName, 3)) = "txt" Then
        TempItem.SubItems(1) = False
    Else
        TempItem.SubItems(1) = True
    End If
    TempI = 1440
    TempI = TempI * 1024
    TempItem.SubItems(2) = TempI
    TempItem.SubItems(4) = comFile.FileName
    TempItem.Text = comFile.FileTitle
End Sub

Private Sub cmdDelete_Click()
    On Error Resume Next
    If lsvQue.SelectedItem.Text = "" Then Exit Sub
    frmAdd.Show
    frmAdd.txtFile.Text = lsvQue.SelectedItem.SubItems(4)
    frmAdd.txtFileSize = lsvQue.SelectedItem.SubItems(2) / 1024
    frmAdd.txtTempName = lsvQue.SelectedItem.SubItems(3)
    frmAdd.IndexNum = lsvQue.SelectedItem.Index
    frmAdd.comFile.FileTitle = lsvQue.SelectedItem.Text
End Sub

Private Sub cmdStart_Click()
    Dim I As Integer
    Dim Temp As String
    Dim TempItem As ListItem
    tmrOverAll.Enabled = True
    cmdStart.Enabled = False
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    Running = True
RemoveItems:
    For I = 1 To lsvQue.ListItems.Count
        If lsvQue.ListItems(I).SubItems(1) = "*****" Then
            lsvQue.ListItems.Remove I
            GoTo RemoveItems
        End If
    Next I
    If chkRestoreAll Then Temp = "echo off" & vbCrLf
    For I = 1 To lsvQue.ListItems.Count
        With lsvQue.ListItems(I)
            lblWorkingOn.Caption = "Working on: " & .Text
            tmrTime.Enabled = True
            SplitFile .SubItems(4), .Text, .SubItems(2), .SubItems(3), _
                .SubItems(1), False, pgbStatus, lblStatus
            tmrTime.Enabled = False
            If chkRestoreAll Then
                Temp = Temp & "call " & .SubItems(4) & "-restore.bat" & vbCrLf
            End If
            .SubItems(1) = "": .SubItems(2) = "": .SubItems(3) = "": .SubItems(4) = ""
            .SubItems(1) = "*****": .SubItems(2) = "Complete in:"
            .SubItems(3) = Int(TimeAmount / 60) & " min(s)"
            .SubItems(4) = TimeAmount Mod 60 & " sec(s)"
            TimeAmount = 0
        End With
    Next I

    If chkRestoreAll Then
        Temp = Temp & "echo on"
        Open App.Path & "\Restore-all.bat" For Output As #1
            Print #1, Temp
        Close #1
    End If
    Running = False
    cmdAdd.Enabled = True
    cmdDelete.Enabled = True
    cmdStart.Enabled = True
    
    tmrOverAll.Enabled = False
    Set TempItem = lsvQue.ListItems.Add()
    TempItem.Text = "----------------------------"
    TempItem.SubItems(1) = "*****"
    TempItem.SubItems(2) = "Overall time:"
    TempItem.SubItems(3) = Int(OverAllTime / 60) & " min(s)"
    TempItem.SubItems(4) = OverAllTime Mod 60 & " sec(s)"
    Set TempItem = Nothing
    OverAllTime = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Running = True Then End
End Sub

Private Sub tmrOverAll_Timer()
    OverAllTime = OverAllTime + 1
End Sub

Private Sub tmrTime_Timer()
    TimeAmount = TimeAmount + 1
End Sub
