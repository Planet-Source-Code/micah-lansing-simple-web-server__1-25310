VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Simple Web Server"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   3270
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Cnt2 
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   3660
      Width           =   1575
   End
   Begin VB.TextBox Cnt1 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3120
      Width           =   1515
   End
   Begin VB.TextBox IName 
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Text            =   "Index.html"
      Top             =   3660
      Width           =   1575
   End
   Begin VB.TextBox Path 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Text            =   "C:\html"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   2745
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   0
      Width           =   3270
   End
   Begin VB.CommandButton Stopcmd 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   4260
      Width           =   855
   End
   Begin VB.CommandButton Listencmd 
      Caption         =   "Listen"
      Height          =   315
      Left            =   420
      TabIndex        =   0
      Top             =   4260
      Width           =   855
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   0
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Label Label4 
      Caption         =   "Connections Made:"
      Height          =   195
      Left            =   1680
      TabIndex        =   10
      Top             =   3420
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Current Connections:"
      Height          =   195
      Left            =   1680
      TabIndex        =   7
      Top             =   2880
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "Index Name"
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   3420
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Path to index"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConnectionCnt As Integer
Dim CurrentConnectionCnt As Integer
Dim Data As String
Dim Newdata As String
Dim Request As String
Private Sub Form_Load()
Winsock1(0).LocalPort = 80
Text1.Text = Winsock1(0).LocalIP
End Sub

Private Sub Listencmd_Click()
On Error GoTo pe
Winsock1(0).Listen
Stopcmd.Enabled = True
Listencmd.Enabled = False
pe:
End Sub

Private Sub Stopcmd_Click()
On Error GoTo pe
Winsock1(0).Close
Listencmd.Enabled = True
Stopcmd.Enabled = False
pe:
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
ConnectionCnt = ConnectionCnt + 1
If ConnectionCnt > 1000 Then ConnectionCnt = 1 'makes sure too many winsock controls arent loaded,
                                               'can be set as small or large as you want(within reason)
Load Winsock1(ConnectionCnt)
Winsock1(ConnectionCnt).Accept requestID
CurrentConnectionCnt = CurrentConnectionCnt + 1
Cnt1.Text = CurrentConnectionCnt
Cnt2.Text = ConnectionCnt
fdsa = Winsock1(ConnectionCnt).RemoteHostIP
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Winsock1(Index).GetData Newdata 'Gets data from remote computer
If Newdata = "" Then Winsock1_SendComplete (Index) 'Fixes error when user refreshes too fast
On Error GoTo pe 'Fixes most errors:)
Request = Mid(Newdata, 5, InStr(5, Newdata, " HTTP/") - 5) 'Gets filename requested by user
If Request <> "/" Then
        Open Path + Request For Binary Access Read As #1
            On Error GoTo 0
            Data = Space(LOF(1))
            Get #1, , Data
            Text1.Text = Data
        Close #1
    If Data = "" Then Kill Path + Request: FourFourError 'if there is no data in file sent "HTTP 404 File not found" error
    'Send Web page, pics text, etc.
    Winsock1(Index).SendData "HTTP/1.0 200 OK" & vbCrLf & "Content-Length: " & Len(Data) & vbCrLf & "Content-Type: text/html" & vbCrLf & vbCrLf & Data
End If
If Request = "/" Then
Open Path & "/" & IName For Binary Access Read As #1
            On Error GoTo 0
            Data = Space(LOF(1))
            Get #1, , Data
            Text1.Text = Data
        Close #1
    If Data = "" Then Kill Path + Request: FourFourError 'if there is no data in file sent "HTTP 404 File not found" error
    'Send Web page, pics text, etc.
    Winsock1(Index).SendData "HTTP/1.0 200 OK" & vbCrLf & "Content-Length: " & Len(Data) & vbCrLf & "Content-Type: text/html" & vbCrLf & vbCrLf & Data
End If
pe:
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1(Index).Close
Unload Winsock1(Index)
CurrentConnectionCnt = CurrentConnectionCnt - 1
Cnt1.Text = CurrentConnectionCnt
End Sub

Private Sub Winsock1_SendComplete(Index As Integer)
Winsock1(Index).Close
Unload Winsock1(Index)
CurrentConnectionCnt = CurrentConnectionCnt - 1
Cnt1.Text = CurrentConnectionCnt
End Sub
Private Sub FourFourError()
Data = "<html><head><title>HTTP 404 File not found.</title></head><body><h1>File not found</h1></body></html>"
End Sub

