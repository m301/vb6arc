VERSION 5.00
Begin VB.Form MAD_USB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MAD USB Manager"
   ClientHeight    =   2415
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   5460
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton command1 
      Caption         =   "Enable USB"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disable USB"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Current Status :"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "©MADSACSoft.com"
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Programmed by : MADhurendra SAChan"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   5175
   End
End
Attribute VB_Name = "MAD_USB"
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
'Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
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
Private Function Registry_Read(Key_Path, Key_Name) As Variant
    On Error Resume Next
    Dim Registry As Object
    Set Registry = CreateObject("WScript.Shell")
    Registry_Read = Registry.RegRead(Key_Path & "\" & Key_Name)
End Function
Private Sub Command1_Click()
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
  MsgBox "All usb devices are enabled.", vbInformation + vbOKOnly, "Done !"
   urefresh
End Sub


Private Sub urefresh()
If (Registry_Read("HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\services\USBSTOR", "Type") = 1) Then
Label4.Caption = "Enabled"
Else
Label4.Caption = "Disabled"
End If


End Sub

 
Private Sub Form_Initialize()
   ' InitCommonControls
End Sub

Private Sub Command2_Click()
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
 MsgBox "All usb devices are disabled. ", vbInformation + vbOKOnly, "Done !"
 urefresh
 End Sub

Private Sub Form_Load()
'MsgBox "Always run this program with administrator rights. OR run this program from 'Run as Administrator' option at context menu or Right Click menu .... To Make Successful Changes!", vbInformation + vbOKOnly, "Information"
urefresh
End Sub

Private Sub Label2_Click()
frmAbout.Show
End Sub

Private Sub Label3_Click()
frmAbout.Show
End Sub

