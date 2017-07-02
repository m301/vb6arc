VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MAD_Process 
   Caption         =   "MAD Process Killer"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11415
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   3000
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   360
      TabIndex        =   1
      Top             =   4080
      Width           =   10455
      Begin VB.CheckBox Check1 
         Caption         =   "Auto Kill New Process"
         Height          =   495
         Left            =   6240
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Open Path of Selected Executable"
         Height          =   495
         Left            =   6240
         TabIndex        =   6
         Top             =   1200
         Width           =   3015
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Make New List"
         Height          =   495
         Left            =   3360
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Kill Selected Process"
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Refresh Process List"
         Height          =   495
         Left            =   3360
         TabIndex        =   3
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Kill All New Process"
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Programmed, Designed And Concept By : MADhurendra SAChan"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2040
         Width           =   7575
      End
      Begin VB.Label Label2 
         Caption         =   "©MADSACSoft.com"
         Height          =   255
         Index           =   1
         Left            =   8160
         TabIndex        =   10
         Top             =   2040
         Width           =   2175
      End
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   4048
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      Caption         =   "MAD Process Killer"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   9
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Process List :"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "MAD_Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iAS, pTot As Integer
Dim pName(100), pName2(100) As String

Private Sub Check1_Click()
If Check1.Value = 1 Then Timer1.Enabled = True Else Timer1.Enabled = False
End Sub

Private Sub Command1_Click()
tKill lvw.SelectedItem.Text
lvw.ListItems.Clear
iAS = 1
GetAllProcesses
MsgBox "Selected process from list is killed ! ", vbInformation + vbOKOnly, "Selected process killed "
End Sub

Private Sub Command2_Click()
lvw.ListItems.Clear
iAS = 1
GetAllProcesses
End Sub

Private Sub Command3_Click()
kill_new
MsgBox "All new process which ran after running this program were killed ! ", vbInformation + vbOKOnly, "All new process killed "
lvw.ListItems.Clear
iAS = 1
GetAllProcesses
End Sub

Private Sub Command4_Click()
pFirstLoad
GetAllProcesses
MsgBox "New process list made ...all program except from current list will be killed when u will click 'Kill all new process' ! ", vbInformation + vbOKOnly, "New process list made ! "

lvw.ListItems.Clear
iAS = 1
GetAllProcesses
End Sub

Private Sub Command5_Click()
Dim arr() As String
On Error GoTo err
If Right(lvw.SelectedItem.SubItems(1), Len(lvw.SelectedItem.SubItems(1)) - 1) <> "" Then
arr() = Split(Right(lvw.SelectedItem.SubItems(1), Len(lvw.SelectedItem.SubItems(1)) - 1), lvw.SelectedItem.SubItems(2))
Shell "explorer " & arr(0)
Else
 MsgBox "Selected process has no path...it can be a system process.... !", vbCritical + vbOKOnly, " Error !"
End If
Exit Sub
err:
MsgBox "An error occured while opening path.... !", vbCritical + vbOKOnly, " Error !"
End Sub

Private Sub Form_Load()
With lvw.ColumnHeaders
.Clear
.Add , , "Name Of Appliation", Me.Width * 0.15
.Add , , "Path Of Application", Me.Width * 0.5
.Add , , "Process Name", Me.Width * 0.1
.Add , , "Process ID", Me.Width * 0.1
End With
pFirstLoad
GetAllProcesses
End Sub
Sub pFirstLoad()
    iAS = 0
    Dim Process As Object
    For Each Process In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process")
        pName2(iAS) = Process.Name
    iAS = iAS + 1
    Next
    pTot = iAS
End Sub
Sub GetAllProcesses()
    iAS = 0
    Dim Process As Object
    For Each Process In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process")
        pName(iAS) = Process.Name
        lvw.ListItems.Add , , Process.Name
        lvw.ListItems(lvw.ListItems.Count).SubItems(1) = " " & Process.ExecutablePath
        lvw.ListItems(lvw.ListItems.Count).SubItems(2) = Process.Caption
        lvw.ListItems(lvw.ListItems.Count).SubItems(3) = Process.ProcessId
    iAS = iAS + 1
    Next
    
End Sub

Sub tKill(tskName As String)
Shell "taskkill /F /IM " & tskName, vbHide
End Sub
Sub kill_new()
Dim pExist As Boolean
Dim i, X As Integer
    iAS = 0
    Dim Process As Object
    For Each Process In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process")
        pName(iAS) = Process.Name
    iAS = iAS + 1
    Next
    
For i = 0 To iAS
    pExist = False
    
    For X = 0 To pTot
        If pName2(X) = pName(i) Then pExist = True
    Next X
    
    If pExist = False Then Shell "taskkill /F /IM " & pName(i), vbHide

Next i

End Sub
Private Sub Form_Resize()

lvw.Width = Me.Width - 350
Frame1.Top = Me.Height - Frame1.Height - 500
lvw.Height = Me.Height - Frame1.Height - 1200
End Sub

Private Sub Label3_Click()
frmAbout.Show
End Sub

Private Sub Timer1_Timer()
kill_new
End Sub
