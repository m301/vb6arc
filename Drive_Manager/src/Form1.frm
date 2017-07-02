VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MAD Drive Manager"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   9015
      Begin VB.Frame Frame8 
         Height          =   3015
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   8175
         Begin VB.CheckBox Check 
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   7560
            TabIndex        =   42
            Top             =   1440
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton Command2 
            Caption         =   "UnHide Selected Drives"
            Height          =   495
            Left            =   4200
            TabIndex        =   7
            Top             =   1680
            Width           =   3015
         End
         Begin VB.CheckBox Check 
            Caption         =   "Z"
            Height          =   255
            Index           =   26
            Left            =   7440
            TabIndex        =   34
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "Y"
            Height          =   255
            Index           =   25
            Left            =   6840
            TabIndex        =   33
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "X"
            Height          =   255
            Index           =   24
            Left            =   6240
            TabIndex        =   32
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "M"
            Height          =   255
            Index           =   13
            Left            =   7440
            TabIndex        =   31
            Top             =   360
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "L"
            Height          =   255
            Index           =   12
            Left            =   6840
            TabIndex        =   30
            Top             =   360
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "K"
            Height          =   255
            Index           =   11
            Left            =   6240
            TabIndex        =   29
            Top             =   360
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "W"
            Height          =   255
            Index           =   23
            Left            =   5520
            TabIndex        =   28
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "V"
            Height          =   255
            Index           =   22
            Left            =   5040
            TabIndex        =   27
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "U"
            Height          =   255
            Index           =   21
            Left            =   4440
            TabIndex        =   26
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "T"
            Height          =   255
            Index           =   20
            Left            =   3840
            TabIndex        =   25
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "S"
            Height          =   255
            Index           =   19
            Left            =   3240
            TabIndex        =   24
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "R"
            Height          =   255
            Index           =   18
            Left            =   2640
            TabIndex        =   23
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "Q"
            Height          =   255
            Index           =   17
            Left            =   2040
            TabIndex        =   22
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "P"
            Height          =   255
            Index           =   16
            Left            =   1440
            TabIndex        =   21
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "O"
            Height          =   255
            Index           =   15
            Left            =   840
            TabIndex        =   20
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "N"
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   19
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "J"
            Height          =   255
            Index           =   10
            Left            =   5640
            TabIndex        =   18
            Top             =   360
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "I"
            Height          =   255
            Index           =   9
            Left            =   5040
            TabIndex        =   17
            Top             =   360
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "H"
            Height          =   255
            Index           =   8
            Left            =   4440
            TabIndex        =   16
            Top             =   360
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "G"
            Height          =   255
            Index           =   7
            Left            =   3840
            TabIndex        =   15
            Top             =   360
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "F"
            Height          =   255
            Index           =   6
            Left            =   3240
            TabIndex        =   14
            Top             =   360
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "E"
            Height          =   255
            Index           =   5
            Left            =   2640
            TabIndex        =   13
            Top             =   360
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "D"
            Height          =   255
            Index           =   4
            Left            =   2040
            TabIndex        =   12
            Top             =   360
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "C"
            Height          =   255
            Index           =   3
            Left            =   1440
            TabIndex        =   11
            Top             =   360
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "B"
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   10
            Top             =   360
            Width           =   495
         End
         Begin VB.CheckBox Check 
            Caption         =   "A"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   495
         End
         Begin VB.CommandButton command1 
            Caption         =   "Hide Selected Drives"
            Height          =   495
            Left            =   1200
            TabIndex        =   8
            Top             =   1680
            Width           =   3015
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Deselect All"
            Height          =   375
            Left            =   4200
            TabIndex        =   5
            Top             =   1200
            Width           =   3015
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Select All"
            Height          =   375
            Left            =   1200
            TabIndex        =   6
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "You Must Restart Your Windows For The Changes To Take Effect"
            Height          =   195
            Left            =   1680
            TabIndex        =   35
            Top             =   2520
            Width           =   4710
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1335
         Left            =   360
         TabIndex        =   1
         Top             =   4320
         Width           =   8175
         Begin VB.CommandButton Command6 
            Caption         =   "Disable USB"
            Height          =   495
            Left            =   4200
            TabIndex        =   2
            Top             =   360
            Width           =   3015
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Enable USB"
            Height          =   495
            Left            =   1200
            TabIndex        =   3
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label5 
            Caption         =   ": No......No need to restart  :-)"
            Height          =   255
            Left            =   4200
            TabIndex        =   41
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Need To Restart?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   -1  'True
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   40
            Top             =   960
            Width           =   3975
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hide Drives In My Computer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   0
         Left            =   2640
         TabIndex        =   37
         Top             =   360
         Width           =   3360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Disable USB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   1
         Left            =   3600
         TabIndex        =   36
         Top             =   3960
         Width           =   1545
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Note : Always run this progaram with administrator rights..."
      Height          =   255
      Left            =   360
      TabIndex        =   43
      Top             =   6840
      Width           =   8895
   End
   Begin VB.Label Label3 
      Caption         =   "Programmed, Designed And Concept By : MADhurendra SAChan"
      Height          =   375
      Left            =   240
      TabIndex        =   39
      Top             =   7320
      Width           =   7335
   End
   Begin VB.Label Label2 
      Caption         =   "©MADSACSoft.com"
      Height          =   255
      Left            =   7680
      TabIndex        =   38
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   0
      Picture         =   "Form1.frx":014A
      Top             =   120
      Width           =   9120
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long

Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_REGISTRY_RECOVERED = 1014&
Const ERROR_REGISTRY_CORRUPT = 1015&
Const ERROR_REGISTRY_IO_FAILED = 1016&
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const REG_SZ = 1
Private Sub Command1_Click()
Dim retvalue As Long
Dim KeyID As Long, keyvalue As Long
Dim subKey As String
Dim regkey As String
Dim a1 As Long
Dim hCurKey As Long
Dim lRegResult As Long
Dim s As String
Dim a As String
If Check(1).Value = 1 Then
    Check(1).Tag = 1
Else
    Check(1).Tag = 0
End If

If Check(2).Value = 1 Then
    Check(2).Tag = 2
Else
    Check(2).Tag = 0
End If
If Check(3).Value = 1 Then
    Check(3).Tag = 4
Else
    Check(3).Tag = 0
End If

If Check(4).Value = 1 Then
    Check(4).Tag = 8
Else
    Check(4).Tag = 0
End If
If Check(5).Value = 1 Then
    Check(5).Tag = 16
Else
    Check(5).Tag = 0
End If

If Check(6).Value = 1 Then
    Check(6).Tag = 32
Else
    Check(6).Tag = 0
End If
If Check(7).Value = 1 Then
    Check(7).Tag = 64
Else
    Check(7).Tag = 0
End If

If Check(8).Value = 1 Then
    Check(8).Tag = 128
Else
    Check(8).Tag = 0
End If
If Check(9).Value = 1 Then
    Check(9).Tag = 256
Else
    Check(9).Tag = 0
End If

If Check(10).Value = 1 Then
    Check(10).Tag = 512
Else
    Check(10).Tag = 0
End If
If Check(11).Value = 1 Then
    Check(11).Tag = 1024
Else
    Check(11).Tag = 0
End If

If Check(12).Value = 1 Then
    Check(12).Tag = 2048
Else
    Check(12).Tag = 0
End If
If Check(13).Value = 1 Then
    Check(13).Tag = 4096
Else
    Check(13).Tag = 0
End If

If Check(14).Value = 1 Then
    Check(14).Tag = 8192
Else
    Check(14).Tag = 0
End If
If Check(15).Value = 1 Then
    Check(15).Tag = 16384
Else
    Check(15).Tag = 0
End If

If Check(16).Value = 1 Then
    Check(16).Tag = 32768
Else
    Check(16).Tag = 0
End If
If Check(17).Value = 1 Then
    Check(17).Tag = 65536
Else
    Check(17).Tag = 0
End If

If Check(18).Value = 1 Then
    Check(18).Tag = 131072
Else
    Check(18).Tag = 0
End If
If Check(19).Value = 1 Then
    Check(19).Tag = 262144
Else
    Check(19).Tag = 0
End If '

If Check(20).Value = 1 Then
    Check(20).Tag = 524288
Else
    Check(20).Tag = 0
End If
If Check(21).Value = 1 Then
    Check(21).Tag = 1048576
Else
    Check(21).Tag = 0
End If

If Check(22).Value = 1 Then
    Check(22).Tag = 2097152
Else
    Check(22).Tag = 0
End If
If Check(23).Value = 1 Then
    Check(23).Tag = 4194304
Else
    Check(23).Tag = 0
End If

If Check(24).Value = 1 Then
    Check(24).Tag = 8388608
Else
    Check(24).Tag = 0
End If
If Check(25).Value = 1 Then
    Check(25).Tag = 16777216
Else
    Check(25).Tag = 0
End If

If Check(26).Value = 1 Then
    Check(26).Tag = 33554432
Else
    Check(26).Tag = 0
End If

Dim x As Integer
a1 = 0
For x = 1 To 26
a1 = a1 + CLng(Check(x).Tag)
Next x

If a1 = 0 Then
    s = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    a = "NoDrives"
    lRegResult = RegOpenKey(HKEY_CURRENT_USER, s, hCurKey)
    lRegResult = RegDeleteValue(hCurKey, a)
    lRegResult = RegCloseKey(hCurKey)
Else
    If a1 <> 0 Then
    regkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    retvalue = RegCreateKey(HKEY_CURRENT_USER, regkey, KeyID)
    subKey = "NoDrives"
    keyvalue = a1
    retvalue = RegSetValueEx(KeyID, subKey, 0&, 4, keyvalue, 4)
End If
End If
MsgBox "Don't forget to restart your computer to Hide drives...", vbInformation + vbOKOnly, "Restart"

End Sub




Private Sub Command2_Click()
Dim hCurKey As Long
Dim lRegResult As Long
Dim s As String
Dim a As String
 s = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
 a = "NoDrives"
 lRegResult = RegOpenKey(HKEY_CURRENT_USER, s, hCurKey)
 lRegResult = RegDeleteValue(hCurKey, a)
 lRegResult = RegCloseKey(hCurKey)
 MsgBox "Don't forget to restart your computer to Show drives...", vbInformation + vbOKOnly, "Restart"

End Sub

Private Sub Command3_Click()
Dim x As Integer
For x = 1 To 26
Check(x).Value = 1
Next x
End Sub

Private Sub Command4_Click()
Dim x As Integer
For x = 1 To 26
Check(x).Value = 0
Next x
End Sub

Private Sub Command5_Click()
Dim retvalue As Long, result As Long
Dim KeyID As Long, keyvalue As Long
Dim subKey As String
Dim bufSize As Long
Dim regkey As String
Dim abc As Long
Dim a1 As Long
Dim hCurKey As Long
Dim lRegResult As Long
Dim s As String
Dim a As String

    regkey = "SYSTEM\ControlSet001\services\USBSTOR"
    retvalue = RegCreateKey(HKEY_LOCAL_MACHINE, regkey, KeyID)
    subKey = "Type"
    keyvalue = "1"
    retvalue = RegSetValueEx(KeyID, subKey, 0&, 4, keyvalue, 4)
MsgBox "All usb devices are Visible....! ", vbInformation + vbOKOnly, "Done !"
    

End Sub

Private Sub Command6_Click()
Dim retvalue As Long, result As Long
Dim KeyID As Long, keyvalue As Long
Dim subKey As String
Dim bufSize As Long
Dim regkey As String
Dim abc As Long
Dim a1 As Long
Dim hCurKey As Long
Dim lRegResult As Long
Dim s As String
Dim a As String
 s = "SYSTEM\ControlSet001\services\USBSTOR"
 a = "Type"
 lRegResult = RegOpenKey(HKEY_LOCAL_MACHINE, s, hCurKey)
 lRegResult = RegDeleteValue(hCurKey, a)
 lRegResult = RegCloseKey(hCurKey)
MsgBox "All usb devices are hidden....Enjoy ! ", vbInformation + vbOKOnly, "Done !"
End Sub

Private Sub Form_Load()
MsgBox "Always run this program with administrator rights. OR run this program from 'Run as Administrator' option at context menu or Right Click menu .... To Make Successful Changes!", vbInformation + vbOKOnly, "Information"
End Sub

Private Sub Label2_Click()
frmAbout.Show
End Sub

Private Sub Label3_Click()
frmAbout.Show
End Sub

