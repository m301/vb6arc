VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form fMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MAD Web RDP Server"
   ClientHeight    =   7575
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   5250
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   14
      Text            =   "http://127.0.0.1"
      Top             =   960
      Width           =   4455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Refresh Details"
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   1320
      Width           =   2295
   End
   Begin VB.HScrollBar arefval 
      Height          =   255
      Left            =   360
      Max             =   600
      Min             =   20
      SmallChange     =   10
      TabIndex        =   11
      Top             =   6240
      Value           =   500
      Width           =   4455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Auto Refresh"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   5400
      Width           =   4335
   End
   Begin VB.Timer aRef 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5280
      Top             =   1080
   End
   Begin VB.HScrollBar pQuality 
      Height          =   255
      Left            =   360
      Max             =   100
      Min             =   10
      TabIndex        =   7
      Top             =   3960
      Value           =   70
      Width           =   4455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop Server"
      Height          =   615
      Left            =   2520
      TabIndex        =   6
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox rdpPort 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Text            =   "1996"
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start Server"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Timer tmrSendData 
      Index           =   0
      Left            =   6240
      Top             =   1560
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   7200
      ScaleHeight     =   315
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh Screenshot"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   4680
      Width           =   4455
   End
   Begin MSWinsockLib.Winsock Sck 
      Index           =   0
      Left            =   6000
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Caption         =   "If nothing appeared ...click here to reset server ..."
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2880
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "©MADSACSoft.com"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   6960
      Width           =   4935
   End
   Begin VB.Label Label5 
      Caption         =   "Programmed, Designed And Concept By : MADhurendra SAChan"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   7200
      Width           =   5055
   End
   Begin VB.Label Label4 
      Caption         =   "Your IP : 127.0.0.1"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Auto Refresh at :"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   5880
      Width           =   4455
   End
   Begin VB.Shape Shape4 
      Height          =   2295
      Left            =   120
      Top             =   4440
      Width           =   4935
   End
   Begin VB.Shape Shape3 
      Height          =   1695
      Left            =   120
      Top             =   240
      Width           =   4935
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   120
      Top             =   1920
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   120
      Top             =   3360
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "Quality Of Screenshot :"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Port :"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblFileProgress 
      AutoSize        =   -1  'True
      Caption         =   "(No Connection)"
      Height          =   195
      Index           =   0
      Left            =   6600
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
      Width           =   1155
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0


'Reserved space around picturebox
Private Const PictureBoxLeft      As Long = 0
Private Const PictureBoxTop       As Long = 0
Private Const PictureBoxRight     As Long = 0
Private Const PictureBoxBottom    As Long = 240   '240 because form has a menu

'Mouse button for grab and drag
Private Const ButtonDrag          As Integer = 1  'Left Mouse
Private PaintLeft           As Long
Private PaintTop            As Long

Private Const TwipsPerPixel       As Long = 15 'Is this ever not true?

Private m_Image                   As New cImage
Private a_Image     As cImage
Private m_Jpeg      As cJpeg
Private m_FileName  As String
Public iMWidth As Integer

' change this to your server name
Private Const ServerName As String = "MAD Cyber Cafe Manager Client Server "

' this project was designed for only one share
' change the path to the directory you want to share
Dim PathShared As String


Private Type ConnectionInfo
    FileNum As Integer  ' file number of the file opened on the current connection
    TotalLength As Long ' total length of data to send (including the header)
    TotalSent As Long   ' total data sent
    FileName As String  ' file name of the file to send
    
    DataStr As String
End Type

Private cInfo() As ConnectionInfo
Private Type RECT
left As Long
top As Long
Right As Long
Bottom As Long
End Type
Private Type PICTDESC
cbSize As Long
pictType As Long
hIcon As Long
hPal As Long
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
(lpPictDesc As PICTDESC, riid As Any, ByVal fOwn As Long, _
IPic As IPicture) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As _
Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, _
ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, _
ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, _
ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, _
ByVal nHeight As Long, ByVal lScreenDC As Long, ByVal xSrc As Long, _
ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
ByVal hDC As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
lpRect As RECT) As Long
' Capture the contents of a window or the entire screen
Function GetScreenSnapshot(Optional ByVal hwnd As Long) As IPictureDisp
Dim targetDC As Long
Dim hDC As Long
Dim tempPict As Long
Dim oldPict As Long
Dim wndWidth As Long
Dim wndHeight As Long
Dim Pic As PICTDESC
Dim rcWindow As RECT
Dim GUID(3) As Long
' provide the right handle for the desktop window
If hwnd = 0 Then hwnd = GetDesktopWindow
' get window's size
GetWindowRect hwnd, rcWindow
wndWidth = rcWindow.Right - rcWindow.left
wndHeight = rcWindow.Bottom - rcWindow.top
' get window's device context
targetDC = GetWindowDC(hwnd)
' create a compatible DC
hDC = CreateCompatibleDC(targetDC)
' create a memory bitmap in the DC just created
' the has the size of the window we're capturing
tempPict = CreateCompatibleBitmap(targetDC, wndWidth, wndHeight)
oldPict = SelectObject(hDC, tempPict)
' copy the screen image into the DC
BitBlt hDC, 0, 0, wndWidth, wndHeight, targetDC, 0, 0, vbSrcCopy
' set the old DC image and release the DC
tempPict = SelectObject(hDC, oldPict)
DeleteDC hDC
ReleaseDC GetDesktopWindow, targetDC
' fill the ScreenPic structure
With Pic
.cbSize = Len(Pic)
.pictType = 1 ' means picture
.hIcon = tempPict
.hPal = 0 ' (you can omit this of course)
End With
' convert the image to a IpictureDisp object
' this is the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
' we use an array of Long to initialize it faster
GUID(0) = &H7BF80980
GUID(1) = &H101ABF32
GUID(2) = &HAA00BB8B
GUID(3) = &HAB0C3000
' create the picture,
' return an object reference right into the function result
OleCreatePictureIndirect Pic, GUID(0), True, GetScreenSnapshot
End Function

Private Sub aRef_Timer()
rRefresh
End Sub

Private Sub arefval_Change()
Label3.Caption = "Auto Refresh at : " & arefval.Value / 10 & " sec."
aRef.Interval = arefval.Value * 10
aRef.Enabled = True
Check1.Value = 1
If arefval.Value / 10 < 30 Then MsgBox " Too less refresh time may slow up your system !", vbCritical + vbOKOnly, "Too  less refresh time !"
End Sub

Private Sub Check1_Click()
If Check1.Value <> 1 Then
aRef.Enabled = False
Else
aRef.Interval = arefval.Value * 10
aRef.Enabled = True
End If
End Sub

Private Sub Command1_Click()
rRefresh
End Sub

Private Sub Command2_Click()
 On Error GoTo err
 rRefresh
 Sck(0).LocalPort = rdpPort.Text ' set this to the port you want the server to listen on...
    Sck(0).Listen
    iMWidth = 650
    DoEvents
    
    If Sck(0).State = sckListening Then lblFileProgress(0).Caption = "00 Listening"
    Command2.Enabled = False
Command3.Enabled = True
Exit Sub

err:
MsgBox "An error occured ....may be port chosen is already in use...try another port !", vbCritical + vbOKOnly, "Error !"

End Sub

Private Sub Command3_Click()
Sck(0).Close
Command2.Enabled = True
Command3.Enabled = False
End Sub

Private Sub Command4_Click()
Label4.Caption = "Your IP : " & Sck(0).LocalIP
Text1.Text = "http://" & Sck(0).LocalIP & ":" & rdpPort.Text
End Sub

Private Sub Form_Load()
PathShared = App.Path & "\"
delTemp
'rRefresh
End Sub
Sub delTemp()
If Dir(PathShared & "snap.bmp") <> "" Then Kill PathShared & "snap.bmp"
If Dir(PathShared & "snap.jpg") <> "" Then Kill PathShared & "snap.jpg"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
delTemp
End Sub

Private Sub Label2_Click(Index As Integer)
frmAbout.Show
End Sub

Private Sub Label5_Click()
frmAbout.Show
End Sub

Private Sub Label6_Click()
On Error GoTo err
PathShared = App.Path & "\"
delTemp
rRefresh
MsgBox "All temperory files were deleted.... if even now the problem is fixed then delete all themperory files from directory containing this application...sorry for trouble...", vbInformation + vbOKOnly, "Reset done !"
err:
End Sub

Private Sub pQuality_Change()
Label2(0).Caption = "Quality Of Screenshot : " & pQuality.Value
End Sub

Private Sub Sck_Close(Index As Integer)
    ' disable the timer (so it does not send more data than neccessary)
    tmrSendData(Index).Enabled = False
    
    ' make sure the connection is closed
    Do
        Sck(Index).Close
        DoEvents
    Loop Until Sck(Index).State = sckClosed
    
    ' clear the info structure
    cInfo(Index).FileNum = 0
    cInfo(Index).FileName = ""
    cInfo(Index).TotalLength = 0
    cInfo(Index).TotalSent = 0
    
    lblFileProgress(Index).Caption = Right("00" & Index, 2) & " Closed"
End Sub

Private Sub Sck_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim K As Integer
    On Error GoTo err
    ' look in the control array for a closed connection
    ' note that it's starting to search at index 1 (not index 0)
    ' since index 0 is the one listening on port 80
    For K = 1 To Sck.UBound
        If Sck(K).State = sckClosed Then Exit For
    Next K
    
    ' if all controls are connected, then create a new one
    If K > Sck.UBound Then
        K = Sck.UBound + 1
        Load Sck(K) ' create a new winsock object
        
        Load lblFileProgress(K) ' load the label to display the progress on each conection
        lblFileProgress(K).top = (lblFileProgress(0).Height + 5) * K
        lblFileProgress(K).Visible = False 'hide_progress
        
        ReDim Preserve cInfo(K) ' create new info structure
        
        Load tmrSendData(K) ' load a new timer for the control
        tmrSendData(K).Enabled = False
        tmrSendData(K).Interval = 1
    End If
    
    ' make sure the info structure contains default values (ie: 0's and "")
    cInfo(K).FileName = ""
    cInfo(0).FileNum = 0
    cInfo(K).TotalLength = 0
    cInfo(K).TotalSent = 0
    
    ' accept the connection on the closed control or the new control
    Sck(K).Accept requestID
err:
End Sub

Private Sub Sck_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim rData As String, sHeader As String, RequestedFile As String, ContentType As String
    Dim CompletePath As String
    On Error GoTo err
    Sck(Index).GetData rData, vbString
    
    If rData Like "GET * HTTP/1.?*" Then
        ' get requested file name
        RequestedFile = LeftRange(rData, "GET ", " HTTP/1.", , ReturnEmptyStr)
        
        ' check if request contains "/../" or "/./" or "*" or "?"
        ' (probably someone trying to get a file that is outside of the share)
        If InStr(1, RequestedFile, "/../") > 0 Or InStr(1, RequestedFile, "/./") > 0 Or _
                InStr(1, RequestedFile, "*") > 0 Or InStr(1, RequestedFile, "?") > 0 Or RequestedFile = "" Then
            
            ' send "Not Found" error ...
            sHeader = "HTTP/1.0 404 Not Found" & vbNewLine & "Server: " & ServerName & vbNewLine & vbNewLine
            cInfo(Index).TotalLength = Len(sHeader)
            Sck(Index).SendData sHeader
        Else
            CompletePath = Replace(PathShared & Replace(RequestedFile, "/", "\"), "\\", "\")
            CompletePath = Replace(CompletePath, "%20", " ")
            Debug.Print CompletePath
            
            If Dir(CompletePath, vbArchive + vbReadOnly + vbDirectory) <> "" Then
                If (GetAttr(CompletePath) And vbDirectory) = vbDirectory Then
                    CompletePath = PathShared & "\snap.jpg"
                     ' requested file exists, open the file, send header, and start the transfer
                    
                    ' display on the label the file name of currect transfer
                    lblFileProgress(Index).Caption = Right("00" & Index, 2) & " Transfering: " & RequestedFile
                    cInfo(Index).FileName = RequestedFile
                    
                    ' since one or more files may be opened at the same time, have to get the free file number
                    cInfo(Index).FileNum = FreeFile
                    Open CompletePath For Binary Access Read As cInfo(Index).FileNum
                    
                    ' get content-type depending on the file extension
                    Select Case LCase(LeftRight(RequestedFile, ".", , ReturnEmptyStr))
                    Case "jpg", "jpeg"
                        ContentType = "Content-Type: image/jpeg"
                    Case "gif"
                        ContentType = "Content-Type: image/gif"
                    Case "htm", "html"
                        ContentType = "Content-Type: text/html"
                    Case "zip"
                        ContentType = "Content-Type: application/zip"
                    Case "mp3"
                        ContentType = "Content-Type: audio/mpeg"
                    Case "m3u", "pls", "xpl"
                        ContentType = "Content-Type: audio/x-mpegurl"
                    Case Else
                        ContentType = "Content-Type: */*"
                    End Select
                    
                    ' build the header
                    sHeader = "HTTP/1.0 200 OK" & vbNewLine & _
                            "Server: " & ServerName & vbNewLine & _
                            ContentType & vbNewLine & _
                            "Content-Length: " & LOF(cInfo(Index).FileNum) & vbNewLine & _
                            vbNewLine
                    
                    ' total data send is the header length + the length of the file requested
                    cInfo(Index).TotalLength = Len(sHeader) + LOF(cInfo(Index).FileNum)
                Else
                    ' requested file exists, open the file, send header, and start the transfer
                    
                    ' display on the label the file name of currect transfer
                    lblFileProgress(Index).Caption = Right("00" & Index, 2) & " Transfering: " & RequestedFile
                    cInfo(Index).FileName = RequestedFile
                    
                    ' since one or more files may be opened at the same time, have to get the free file number
                    cInfo(Index).FileNum = FreeFile
                    Open CompletePath For Binary Access Read As cInfo(Index).FileNum
                    
                    ' get content-type depending on the file extension
                    Select Case LCase(LeftRight(RequestedFile, ".", , ReturnEmptyStr))
                    Case "jpg", "jpeg"
                        ContentType = "Content-Type: image/jpeg"
                    Case "gif"
                        ContentType = "Content-Type: image/gif"
                    Case "htm", "html"
                        ContentType = "Content-Type: text/html"
                    Case "zip"
                        ContentType = "Content-Type: application/zip"
                    Case "mp3"
                        ContentType = "Content-Type: audio/mpeg"
                    Case "m3u", "pls", "xpl"
                        ContentType = "Content-Type: audio/x-mpegurl"
                    Case Else
                        ContentType = "Content-Type: */*"
                    End Select
                    
                    ' build the header
                    sHeader = "HTTP/1.0 200 OK" & vbNewLine & _
                            "Server: " & ServerName & vbNewLine & _
                            ContentType & vbNewLine & _
                            "Content-Length: " & LOF(cInfo(Index).FileNum) & vbNewLine & _
                            vbNewLine
                    
                    ' total data send is the header length + the length of the file requested
                    cInfo(Index).TotalLength = Len(sHeader) + LOF(cInfo(Index).FileNum)
                End If
                
                ' send the header, the Sck_SendComplete event is gonna send the file...
                Sck(Index).SendData sHeader
            Else
                ' send "Not Found" if file does not exsist on the share
                sHeader = "HTTP/1.0 404 Not Found" & vbNewLine & "Server: " & ServerName & vbNewLine & vbNewLine
                cInfo(Index).TotalLength = Len(sHeader)
                Sck(Index).SendData sHeader
            End If
        End If
    Else
        ' sometimes the browser makes "HEAD" requests (but it's not inplemented in this project)
        sHeader = "HTTP/1.0 501 Not Implemented" & vbNewLine & "Server: " & ServerName & vbNewLine & vbNewLine
        cInfo(Index).TotalLength = Len(sHeader)
        Sck(Index).SendData sHeader
    End If
err:
End Sub

Private Function BuildHTMLDirList(ByVal Root As String, ByVal DirToList As String)
    Dim Dirs As New Collection, Files As New Collection
    Dim sDir As String, Path As String, HTML As String, K As Long
    
    Root = Replace(Root, "/", "\")
    DirToList = Replace(DirToList, "/", "\")
    
    If Right(Root, 1) <> "\" Then Root = Root & "\"
    If left(DirToList, 1) = "\" Then DirToList = Mid(DirToList, 2)
    If Right(DirToList, 1) <> "\" Then DirToList = DirToList & "\"
    
    DirToList = Replace(DirToList, "%20", " ")
    
    sDir = Dir(Replace(Root & DirToList, "\\", "\") & "*.*", vbArchive + vbDirectory + vbReadOnly)
    
    Do Until Len(sDir) = 0
        If sDir <> ".." And sDir <> "." Then
            Path = Replace(Root & DirToList, "\\", "\") & sDir
            
            
            If (GetAttr(Path) And vbDirectory) = vbDirectory Then
                Dirs.Add sDir
            Else
                Files.Add sDir
            End If
        End If
        
        sDir = Dir
    Loop
    
    HTML = "<html><body>"
    
    If Dirs.Count > 0 Then
        HTML = HTML & "<b>Directories:</b><br>"
        
        For K = 1 To Dirs.Count
            HTML = HTML & "<a href=""" & _
                Replace(Replace("/" & DirToList & Dirs(K), "\", "/"), "//", "/") & """>" & _
                Dirs(K) & "</a><br>" & vbNewLine
        Next K
    End If
    
    If Files.Count > 0 Then
        HTML = HTML & "<br><b>Files:</b><br><table width=""100%"" border=""1"" cellpadding=""3"" cellspacing=""2"">" & vbNewLine
        
        For K = 1 To Files.Count
            HTML = HTML & "<tr>" & vbNewLine
            HTML = HTML & "<td width=""100%""><a href=""" & _
                Replace(Replace("/" & DirToList & Files(K), "\", "/"), "//", "/") & """>" & _
                Files(K) & "</a></td>" & vbNewLine
            
            HTML = HTML & "<td nowrap>" & _
                Format(FileLen(Replace(Root & DirToList, "\\", "\") & Files(K)) / 1024#, "###,###,###,##0") & _
                " KBytes</td>" & vbNewLine
            HTML = HTML & "</tr>" & vbNewLine
        Next K
        
        HTML = HTML & "</table>" & vbNewLine
    End If
    
    If Dirs.Count = 0 And Files.Count = 0 Then
        HTML = HTML & "This folder is empty."
    End If
    
    BuildHTMLDirList = HTML & "</body></html>"
End Function

Private Sub Sck_SendComplete(Index As Integer)
    If cInfo(Index).TotalSent >= cInfo(Index).TotalLength Then
        ' if all data was sent, then close the connection
        Sck_Close Index
    Else
        ' still have data to send, let the timer do that...
        
        ' if you want to slow down the connection set the interval to a higher values
        ' right now it's as fast as it can be
        tmrSendData(Index).Interval = 1
        tmrSendData(Index).Enabled = True ' start the timer
    End If
End Sub

Private Sub Sck_SendProgress(Index As Integer, ByVal BytesSent As Long, ByVal bytesRemaining As Long)
    ' keep track of how much data was sent
    cInfo(Index).TotalSent = cInfo(Index).TotalSent + BytesSent
    
    ' display file progress
    If cInfo(Index).FileNum > 0 Then
        On Error Resume Next
        
        lblFileProgress(Index).Caption = Right("00" & Index, 2) & " Transfering: " & cInfo(Index).FileName & "   - " & _
            cInfo(Index).TotalSent & " of " & LOF(cInfo(Index).FileNum) & " bytes sent, " & _
            Format(cInfo(Index).TotalSent / LOF(cInfo(Index).FileNum) * 100#, "00.00") & " %Done."
        
        ' if file size if 0 length, it gives a "division by 0" error, so just clear it
        If err.Number <> 0 Then err.Clear
    End If
End Sub

Private Sub tmrSendData_Timer(Index As Integer)
   On Error GoTo 100
   ' send 2KBytes at one time, then wait until it's sent, then send the other 2KBytes and so on
    Const BufferLength As Long = 1024 * 2
    Dim Buffer As String
    
    If cInfo(Index).FileNum = -1 Then ' send data from a string instead of a file on the hard drive
        Buffer = left(cInfo(Index).DataStr, BufferLength)
        cInfo(Index).DataStr = Mid(cInfo(Index).DataStr, BufferLength + 1)
        
        Sck(Index).SendData Buffer ' send the data on the current connection
        
        If Len(cInfo(Index).DataStr) = 0 Then cInfo(Index).FileNum = 0
        
    ElseIf cInfo(Index).FileNum > 0 Then ' do this code ONLY if a file was opened
        If Loc(cInfo(Index).FileNum) + BufferLength > LOF(cInfo(Index).FileNum) Then
            ' the remaining data is less than the buffer length, so load only the remaining data
            
            Buffer = String(LOF(cInfo(Index).FileNum) - Loc(cInfo(Index).FileNum), 0)
        Else
            Buffer = String(BufferLength, 0)
        End If
        
        Get cInfo(Index).FileNum, , Buffer ' get data from file
        
        Sck(Index).SendData Buffer ' send the data on the current connection
        
        If Loc(cInfo(Index).FileNum) >= LOF(cInfo(Index).FileNum) Then
            ' no data remaining to send
            Close cInfo(Index).FileNum
            
            ' if file is closed, set filenumber to 0, in case the timer get's called again don't send any more data
            cInfo(Index).FileNum = 0
        End If
    End If
    
    ' Sck_SendComplete event will enable the time back when current data is sent
    tmrSendData(Index).Enabled = False
    Exit Sub
100
Sck(Index).Close

End Sub


Public Function rRefresh()
On Error GoTo err
Dim std As StdPicture
Dim MyPic As StdPicture
delTemp

SavePicture GetScreenSnapshot, PathShared & "snap.bmp"
    Dim FileName As String

'FileName = FileDialog(Me, False, "Open Picture File", "Picture Files|*.jpg;*.jpeg;*.gif;*.bmp;*.wmp;*.rle;*.cur;*.ico;*.emf|All Files [*.*]|*.*")
 
FileName = PathShared & "snap.bmp"
        Set MyPic = LoadPicture(FileName)
            Set m_Image = New cImage
            m_Image.CopyStdPicture MyPic
        Set MyPic = Nothing

SaveImage m_Image, PathShared & "snap.jpg"
i_Save
err:
End Function
Public Function i_Save()
On Error GoTo err
   Set m_Jpeg = New cJpeg
    'cboSubSample.ListIndex = 3

     m_Jpeg.Quality = pQuality.Value

       'Sample the cImage by hDC
        m_Jpeg.SampleHDC a_Image.hDC, a_Image.Width, a_Image.Height

       'Delete file if it exists
        RidFile m_FileName

       'Save the JPG file
        m_Jpeg.SaveFile m_FileName
    
    Set a_Image = Nothing
    Set m_Jpeg = Nothing
err:
End Function


Public Sub SaveImage(TheImage As cImage, FileName As String)
    Set a_Image = TheImage 'Call this before the form loads to initialize it
    m_FileName = FileName
End Sub


