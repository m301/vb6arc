VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MAD HTTP Server"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtIP 
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   4935
      Begin VB.CheckBox Check2 
         Caption         =   "Enable Directory Browsing "
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Use ""index.html"" as homepage"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   3495
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start Server"
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox rdpPort 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Text            =   "1996"
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop Server"
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Refresh Details"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Text            =   "http://127.0.0.1"
      Top             =   960
      Width           =   4455
   End
   Begin VB.Timer tmrSendData 
      Index           =   0
      Left            =   7425
      Top             =   90
   End
   Begin MSWinsockLib.Winsock Sck 
      Index           =   0
      Left            =   6930
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Server Root Directory"
      Height          =   1935
      Left            =   480
      TabIndex        =   11
      Top             =   2040
      Width           =   4455
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4215
      End
      Begin VB.DirListBox Dir1 
         Height          =   1215
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   4215
      End
   End
   Begin VB.Label Label2 
      Caption         =   "©MADSACSoft.com"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   6840
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Port :"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
   Begin VB.Shape Shape2 
      Height          =   3015
      Left            =   240
      Top             =   1920
      Width           =   4935
   End
   Begin VB.Shape Shape3 
      Height          =   1695
      Left            =   240
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label Label4 
      Caption         =   "Your IP : 127.0.0.1"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label lblFileProgress 
      AutoSize        =   -1  'True
      Caption         =   "(No Connection)"
      Height          =   195
      Index           =   0
      Left            =   5280
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1155
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' change this to your server name
Private Const ServerName As String = "MAD HTTP Server"

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

Private CInfo() As ConnectionInfo

Private Sub Command2_Click()
 On Error GoTo err
PathShared = Dir1.Path
 Sck(0).LocalPort = rdpPort.Text ' set this to the port you want the server to listen on...
    'Sck(0).Listen
    Sck(0).Bind rdpPort.Text, txtIP.Text
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

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
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
    CInfo(Index).FileNum = 0
    CInfo(Index).FileName = ""
    CInfo(Index).TotalLength = 0
    CInfo(Index).TotalSent = 0
    
    lblFileProgress(Index).Caption = Right("00" & Index, 2) & " Closed"
End Sub

Private Sub Sck_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim K As Integer
    
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
        lblFileProgress(K).Top = (lblFileProgress(0).Height + 5) * K
        lblFileProgress(K).Visible = False
        
        ReDim Preserve CInfo(K) ' create new info structure
        
        Load tmrSendData(K) ' load a new timer for the control
        tmrSendData(K).Enabled = False
        tmrSendData(K).Interval = 1
    End If
    
    ' make sure the info structure contains default values (ie: 0's and "")
    CInfo(K).FileName = ""
    CInfo(0).FileNum = 0
    CInfo(K).TotalLength = 0
    CInfo(K).TotalSent = 0
    
    ' accept the connection on the closed control or the new control
    Sck(K).Accept requestID
End Sub

Private Sub Sck_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim rData As String, sHeader As String, RequestedFile As String, ContentType As String
    Dim CompletePath As String
    
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
            CInfo(Index).TotalLength = Len(sHeader)
            Sck(Index).SendData sHeader
        Else
            CompletePath = Replace(PathShared & Replace(RequestedFile, "/", "\"), "\\", "\")
            CompletePath = Replace(CompletePath, "%20", " ")
            Debug.Print CompletePath
            
            If Dir(CompletePath, vbArchive + vbReadOnly + vbDirectory) <> "" Then
                If (GetAttr(CompletePath) And vbDirectory) = vbDirectory Then
                    ' the request if for a directory listing...
                    If Check1.Value <> 1 And Check2.Value = 1 Then
                    
                    
                    CInfo(Index).DataStr = BuildHTMLDirList(PathShared, RequestedFile)
                    CInfo(Index).FileNum = -1
                    ' build the header
                    sHeader = "HTTP/1.0 200 OK" & vbNewLine & _
                            "Server: " & ServerName & vbNewLine & _
                            "Content-Type: text/html" & vbNewLine & _
                            "Content-Length: " & Len(CInfo(Index).DataStr) & vbNewLine & _
                            vbNewLine
                    
                    ' total data send is the header length + the length of the file requested
                    CInfo(Index).TotalLength = Len(sHeader) + Len(CInfo(Index).DataStr)
                    
                    
                    ElseIf Check1.Value = 1 Then
                     lblFileProgress(Index).Caption = Right("00" & Index, 2) & " Transfering: index.html"
                    CInfo(Index).FileName = "index.html"
                    
                    ' since one or more files may be opened at the same time, have to get the free file number
                    CInfo(Index).FileNum = FreeFile
                    Open App.Path & "\index.html" For Binary Access Read As CInfo(Index).FileNum
                    ContentType = "Content-Type: text/html"
                

                    sHeader = "HTTP/1.0 200 OK" & vbNewLine & _
                            "Server: " & ServerName & vbNewLine & _
                            ContentType & vbNewLine & _
                            "Content-Length: " & LOF(CInfo(Index).FileNum) & vbNewLine & _
                            vbNewLine
                    
                    ' total data send is the header length + the length of the file requested
                    CInfo(Index).TotalLength = Len(sHeader) + LOF(CInfo(Index).FileNum)
                    Else
                    CInfo(Index).DataStr = "<html><title>" & ServerName & "</title><body><h1 align=center>" & ServerName & " Configured successfully !</h1><br>Check setting if this appears you as an error !</body></html>"
                    CInfo(Index).FileNum = -1
                    ' build the header
                    sHeader = "HTTP/1.0 200 OK" & vbNewLine & _
                            "Server: " & ServerName & vbNewLine & _
                            "Content-Type: text/html" & vbNewLine & _
                            "Content-Length: " & Len(CInfo(Index).DataStr) & vbNewLine & _
                            vbNewLine
                    
                    ' total data send is the header length + the length of the file requested
                    CInfo(Index).TotalLength = Len(sHeader) + Len(CInfo(Index).DataStr)
                    End If
                    
                Else
                    ' requested file exists, open the file, send header, and start the transfer
                    
                    ' display on the label the file name of currect transfer
                    lblFileProgress(Index).Caption = Right("00" & Index, 2) & " Transfering: " & RequestedFile
                    CInfo(Index).FileName = RequestedFile
                    
                    ' since one or more files may be opened at the same time, have to get the free file number
                    CInfo(Index).FileNum = FreeFile
                    Open CompletePath For Binary Access Read As CInfo(Index).FileNum
                    
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
                            "Content-Length: " & LOF(CInfo(Index).FileNum) & vbNewLine & _
                            vbNewLine
                    
                    ' total data send is the header length + the length of the file requested
                    CInfo(Index).TotalLength = Len(sHeader) + LOF(CInfo(Index).FileNum)
                End If
                
                ' send the header, the Sck_SendComplete event is gonna send the file...
                Sck(Index).SendData sHeader
            Else
                ' send "Not Found" if file does not exsist on the share
                sHeader = "HTTP/1.0 404 Not Found" & vbNewLine & "Server: " & ServerName & vbNewLine & vbNewLine
                CInfo(Index).TotalLength = Len(sHeader)
                Sck(Index).SendData sHeader
            End If
        End If
    Else
        ' sometimes the browser makes "HEAD" requests (but it's not inplemented in this project)
        sHeader = "HTTP/1.0 501 Not Implemented" & vbNewLine & "Server: " & ServerName & vbNewLine & vbNewLine
        CInfo(Index).TotalLength = Len(sHeader)
        Sck(Index).SendData sHeader
    End If
End Sub

Private Function BuildHTMLDirList(ByVal Root As String, ByVal DirToList As String)
    Dim Dirs As New Collection, Files As New Collection
    Dim sDir As String, Path As String, HTML As String, K As Long
    
    Root = Replace(Root, "/", "\")
    DirToList = Replace(DirToList, "/", "\")
    
    If Right(Root, 1) <> "\" Then Root = Root & "\"
    If Left(DirToList, 1) = "\" Then DirToList = Mid(DirToList, 2)
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
    
    HTML = "<html><body><h1 align=center>" & ServerName & "</h1>"
    
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
    If CInfo(Index).TotalSent >= CInfo(Index).TotalLength Then
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

Private Sub Sck_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    ' keep track of how much data was sent
    CInfo(Index).TotalSent = CInfo(Index).TotalSent + bytesSent
    
    ' display file progress
    If CInfo(Index).FileNum > 0 Then
        On Error Resume Next
        
        lblFileProgress(Index).Caption = Right("00" & Index, 2) & " Transfering: " & CInfo(Index).FileName & "   - " & _
            CInfo(Index).TotalSent & " of " & LOF(CInfo(Index).FileNum) & " bytes sent, " & _
            Format(CInfo(Index).TotalSent / LOF(CInfo(Index).FileNum) * 100#, "00.00") & " %Done."
        
        ' if file size if 0 length, it gives a "division by 0" error, so just clear it
        If err.Number <> 0 Then err.Clear
    End If
End Sub

Private Sub tmrSendData_Timer(Index As Integer)
    ' send 2KBytes at one time, then wait until it's sent, then send the other 2KBytes and so on
    Const BufferLength As Long = 1024 * 2
    Dim Buffer As String
    
    If CInfo(Index).FileNum = -1 Then ' send data from a string instead of a file on the hard drive
        Buffer = Left(CInfo(Index).DataStr, BufferLength)
        CInfo(Index).DataStr = Mid(CInfo(Index).DataStr, BufferLength + 1)
        
        Sck(Index).SendData Buffer ' send the data on the current connection
        
        If Len(CInfo(Index).DataStr) = 0 Then CInfo(Index).FileNum = 0
        
    ElseIf CInfo(Index).FileNum > 0 Then ' do this code ONLY if a file was opened
        If Loc(CInfo(Index).FileNum) + BufferLength > LOF(CInfo(Index).FileNum) Then
            ' the remaining data is less than the buffer length, so load only the remaining data
            
            Buffer = String(LOF(CInfo(Index).FileNum) - Loc(CInfo(Index).FileNum), 0)
        Else
            Buffer = String(BufferLength, 0)
        End If
        
        Get CInfo(Index).FileNum, , Buffer ' get data from file
        Sck(Index).SendData Buffer ' send the data on the current connection
        
        If Loc(CInfo(Index).FileNum) >= LOF(CInfo(Index).FileNum) Then
            ' no data remaining to send
            Close CInfo(Index).FileNum
            
            ' if file is closed, set filenumber to 0, in case the timer get's called again don't send any more data
            CInfo(Index).FileNum = 0
        End If
    End If
    
    ' Sck_SendComplete event will enable the time back when current data is sent
    tmrSendData(Index).Enabled = False
End Sub


