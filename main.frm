VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Main 
   Caption         =   "Elucid Software Winsock Downloader"
   ClientHeight    =   5220
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   0
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Timer tmrUpdateProgress 
      Interval        =   1
      Left            =   0
      Top             =   1920
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   5055
      Begin VB.TextBox txtHead 
         Appearance      =   0  'Flat
         Height          =   2295
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Text            =   "http://tucows.erols.com/files4/bzfinst.exe"
      Top             =   240
      Width           =   5055
   End
   Begin VB.Timer tmrTimeLeft 
      Interval        =   1000
      Left            =   0
      Top             =   1560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Download"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "&File Download Progress"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   5055
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C00000&
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   4785
         TabIndex        =   2
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label lblSize 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblRecieve 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblSpeed 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   15
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblElapsed 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblRemaining 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Elapsed Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   12
         Top             =   960
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Time Remaining:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Speed:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3720
         TabIndex        =   10
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Recieved Size:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   9
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Size:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   750
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Stop"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Pause"
      Height          =   375
      Left            =   2880
      TabIndex        =   18
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label txtResume 
      Height          =   255
      Left            =   2880
      TabIndex        =   21
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter the url in which the file is located:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2835
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu showheader 
         Caption         =   "&Show Header"
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu aboutdownloader 
         Caption         =   "&About Downloader"
      End
      Begin VB.Menu ElucidOnWeb 
         Caption         =   "&Elucid Software Webpage"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DATA As String
Dim Percent%
Dim BeginTransfer As Single
Dim BytesAlreadySent As Single
Dim BytesRemaining As Single
Dim Header As Variant
Dim Status As String
Dim TransferRate As Single
'Dim TimeLeft As String
'Dim TimerVal As Single


Function ConvertTime(TheTime As Single)
    Dim NewTime As String
    Dim Sec As Single
    Dim Min As Single
    Dim H As Single

    If TheTime > 60 Then
        Sec = TheTime
        Min = Sec / 60
        Min = Int(Min)
        Sec = Sec - Min * 60
        H = Int(Min / 60)
        Min = Min - H * 60
        NewTime = H & ":" & Min & ":" & Sec
        If H < 0 Then H = 0
        If Min < 0 Then Min = 0
        If Sec < 0 Then Sec = 0
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If


    If TheTime < 60 Then
        NewTime = "00:00:" & TheTime
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
End Function
Public Function StartUpdate(strURL As String)
BytesAlreadySent = 1
If strURL = "" Then Exit Function
URL = strURL
Dim Pos%, LENGTH%, NextPos%, LENGTH2%, POS2%, POS3%
    Pos = InStr(strURL, "://") 'Record position of ://
    LENGTH2 = Len("://") 'Record the length of it
    LENGTH = Len(strURL) 'Length of the entire url
        If InStr(strURL, "://") Then  ' check if they entered the http:// or ftp://
        strURL = Right(strURL, LENGTH - LENGTH2 - Pos + 1) ' remove http:// or ftp://
        End If
            If InStr(strURL, "/") Then 'looks for the first / mark going from left to right
            POS2 = InStr(strURL, "/") 'gets the position of the / mark
'-----------------GET THE FILENAME-------------
            Dim StrFile$: StrFile = strURL 'load the variables into each other
            Do Until InStr(StrFile, "/") = 0 'Do the loop until all is left is the filename
            LENGTH2 = Len(StrFile) 'get the length of the filename every time its passed over by the loop
            POS3 = InStr(StrFile, "/") 'find the / mark
            StrFile = Right(strURL, LENGTH2 - POS3) 'slash it down removing everything before the / mark including the / mark...
            Loop
            FileName = StrFile
'----------------END GET FILE NAME--------------
            strSvrURL = Left(strURL, POS2 - 1) 'removes everything after the / mark leaving just the server name as the end result
            End If
'-----------END TRIM THE URL FOR THE SERVER NAME-----------

End Function
Public Sub Reset()
CloseSocket
DATA = ""
Percent = 0
BeginTransfer = 0
BytesAlreadySent = 1
BytesRemaining = 0
Status = ""
Header = ""
RESUMEFILE = False
UpdateProgress Picture1, 0
Command1.Enabled = True
End Sub
Public Sub CloseSocket()
Do Until Winsock.State = 0
Winsock.Close
Winsock.LocalPort = 0
Close #1
Loop
End Sub

Private Sub aboutdownloader_Click()
frmAbout.Show
End Sub

Private Sub cmdRun_Click()
OpenIt Me, FilePathName
End Sub

Private Sub Command1_Click()
StartUpdate Text1
frmSave.Show
lblStatus.Visible = False
Picture1.Visible = True
End Sub

Private Sub Command2_Click()
If BytesRemaining > BytesAlreadySent Then
If Winsock.State > 0 Then
DATA = ""
BeginTransfer = 0
Status = ""
Header = ""
CloseSocket
Picture1.Visible = False
lblStatus.Visible = True
lblStatus.Caption = "Download Paused"
Else
Picture1.Visible = True
lblStatus.Visible = False
FileLength = FileLen(FilePathName)
RESUMEFILE = True
Main.Winsock.Connect strSvrURL, 80
End If

End If
End Sub

Private Sub Command3_Click()
If Winsock.State > 0 Then
CloseSocket
MsgBox "Transfer Aborted!", vbExclamation, "Aborted"
Reset
End If
End Sub

Private Sub ElucidOnWeb_Click()
OpenIt Me, "http://elucidsoftware.hypermart.net"
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Unix = False
Me.Height = 3150
RESUMEFILE = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
CloseSocket
End Sub

Private Sub Form_Unload(Cancel As Integer)
CloseSocket
End Sub
Private Sub Progtmr_Timer()
End Sub

Private Sub showheader_Click()
If Me.Height = 5940 Then
Me.Height = 3150
Else
Me.Height = 5940
End If
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub tmrTimeLeft_Timer()
'On Error Resume Next
If BytesRemaining > 0 And BytesAlreadySent > 0 Then
If BytesRemaining <= BytesAlreadySent Then
lblSpeed = 0
CloseSocket
lblElapsed = Format(Hr & ":" & Min & ":" & Sec, "HH:MM:SS")
Command1.Enabled = False
cmdRun.Enabled = True
Picture1.Visible = False
lblStatus.Visible = True
lblStatus.Caption = "Download Completed"
Reset
Else
    Sec = Sec + 1
    If Sec >= 60 Then
    Sec = 0
    Min = Min + 1
    ElseIf Min >= 60 Then
    Min = 0
    Hr = Hr + 1
    End If
Command1.Enabled = True
cmdRun.Enabled = False
lblElapsed = Format(Hr & ":" & Min & ":" & Sec, "HH:MM:SS")
'The reason I divide the difference of bytesalreadysent and bytesremaining is becuase they are in bytes right now.. I want it to be in KB so it can be Kbps and not bps
lblRemaining = ConvertTime(Int(((BytesRemaining - BytesAlreadySent) / 1024) / TransferRate))
lblSpeed = TransferRate
End If

End If
End Sub

Private Sub tmrUpdateProgress_Timer()
On Error Resume Next
If BytesAlreadySent > 0 And BytesRemaining > 0 Then
lblRecieve = File_ByteConversion(BytesAlreadySent)
lblSize = File_ByteConversion(BytesRemaining)
Percent = Format((BytesAlreadySent / BytesRemaining) * 100, "00") 'calculates the percentage completed
UpdateProgress Picture1, Percent 'updates progress bar with new percentage rate
End If
End Sub

Private Sub Winsock_Close()
FormsOnTop Me, False
End Sub

Private Sub Winsock_Connect()
Dim strCommand As String
 
 On Error Resume Next
 
 'Fixed the bug with unix servers thanks to Michael Pauletta
 
 If Not Unix Then
  strCommand = "GET " + URL + " HTTP/1.0" + vbCrLf 'tells server to GET the file if you just want the header info and not the data change "GET " to "HEAD
 Else
    strCommand = "GET " + "/" + FileName + " HTTP/1.0" + vbCrLf 'tells server to GET the file if you just want the header info and not the data change "GET "to "HEAD "
 End If
 
     strCommand = strCommand + "Accept: *.*, */*" + vbCrLf
 If RESUMEFILE = True Then strCommand = strCommand + "Range: bytes=" & FileLength & "-" & vbCrLf
    strCommand = strCommand + "User-Agent: Conquest" & vbCrLf
 
 
 If Not Unix Then
    strCommand = strCommand + "Referer: " & strSvrURL & vbCrLf
 Else
    strCommand = strCommand + "Host: " & strSvrURL & vbCrLf
 End If
 
 
    strCommand = strCommand + vbCrLf
    Winsock.SendData strCommand 'sends a header to the server instructing it what to do!
    BeginTransfer = Timer 'start timer for transfer rate

End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
Winsock.GetData DATA, vbString
If InStr(DATA, "Content-Type:") Then 'find out if this chunk has the header..you can change that to anything that the header contains
        
        If RESUMEFILE = True Then 'check to see if its gonna resume ok or not..This is actually the worst way to check this.
            If InStr(DATA, "HTTP/1.1 206 Partial Content") = 0 Then
            MsgBox "Server did not accept resuming.", vbCritical, "No Resuming Support"
            Exit Sub
            Reset
            CloseSocket
            End If
        End If
        
    If InStr(DATA, "404 Not Found") > 0 Then
   If Not Unix Then
    Unix = True
    Reset
    CloseSocket
    Main.Winsock.Connect strSvrURL, 80
    Exit Sub
   End If
   Unix = False
   MsgBox "File not found on this server.", vbCritical, "File Not Found"
   Reset
   CloseSocket
   Exit Sub
    End If
        
    Dim Pos%, LENGTH%, HEAD$
    Pos = InStr(DATA, vbCrLf & vbCrLf) ' find out where the header and the data is split apart
    LENGTH = Len(DATA) 'get the length of the data chunk
    HEAD = Left(DATA, Pos - 1) 'Get the header from the chunk of data and ignore the data content
    DATA = Right(DATA, LENGTH - Pos - 3) 'Get the data from the first chunk that contains the header also
    Header = Header & HEAD 'Append the header to header text box

If RESUMEFILE = True Then
BytesAlreadySent = FileLength + 1
BytesRemaining = GETDATAHEAD(Header, "Content-Length:")
BytesRemaining = BytesRemaining + FileLength
Else
BytesRemaining = GETDATAHEAD(Header, "Content-Length:")
End If
txtHead = Header
End If

'-----------BEGIN WRITE CHUNK TO FILE CODE--------
        Open FilePathName For Binary Access Write As #1 'opens file for output
        Put #1, BytesAlreadySent, DATA 'writes data to the end of file
        BytesAlreadySent = Seek(1)
        Close #1 'close file for now until next data chunk is available
'--------------------------------------------------

'Lets explain this a bit..The variable BeginTransfer is given the starting value of the
'timer which in case you dont know is the amount of seconds til midnight but that has
'nothing to do with this. Anyways so its given the amount for the start time and then
'when this event below is fired for the first time the timer will be given the value again
'since your system clock was ticking along while the operation between the two of these
'events happened the number will be different.  The two values are subtracted and divided
'by the amount recieved and then by 1000 and put into a readable format
If RESUMEFILE = False Then
'This is pretty straightforward if you ever taken math before you can tell what im doing!
TransferRate = Format(Int(BytesAlreadySent / (Timer - BeginTransfer)) / 1000, "####.00")
Else
'If you dont subtract the difference you will get a really large and odd download speed hehe.
TransferRate = Format(Int((BytesAlreadySent - FileLength) / (Timer - BeginTransfer)) / 1000, "####.00")
End If
End Sub

