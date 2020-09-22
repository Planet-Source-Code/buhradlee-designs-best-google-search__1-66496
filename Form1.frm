VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1920
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "google.com"
      RemotePort      =   80
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      Height          =   1695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "HAI"
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'first off i dont comment alot so dont expect to explain this.
'Credits go to NightQuest And Sniper of irc.hackthissite.org channel is #vb
'you better vote for this, we worked our butts off to accomplish this
' we used regex and winsock as you can tell
'if we dont get an award for this i swear ill never make another submission
' to pscode again
' so please vote for this
' its the best google search youll ever see.

Public Function stripshit(X As String)
X = Replace(X, "<b>", "")
X = Replace(X, "</b>", "")
stripshit = X
End Function
Public Function getcontent(ByVal Html As String) As String
Dim sc As CStrCat
Dim m As Match
Dim regex As RegExp
Set sc = New CStrCat
Set regex = New RegExp
regex.IgnoreCase = True
regex.Global = True
regex.Pattern = "<a title=\x22[^>]*\x22 href=([^>]*)>(.*?)</a>"
sc.MaxLength = Len(Html)
For Each m In regex.Execute(Html)
    sc.AddStr stripshit(m.SubMatches(1)) & " :: " & stripshit(m.SubMatches(0)) & vbCrLf
Next
getcontent = sc
End Function
Private Sub Command1_Click()
  If (Winsock1.State <> sckClosed) Then Winsock1.Close
  Text2 = ""
  Winsock1.Connect
End Sub
Private Sub Winsock1_Close()
  If (Winsock1.State <> sckClosed) Then Winsock1.Close
End Sub
Private Sub Winsock1_Connect()
  Winsock1.SendData _
  "GET /ie?q=" & Replace(Text1, " ", "+") & " HTTP/1.1" & vbCrLf & _
  "Host: www.google.com" & vbCrLf & _
  "User-Agent: Mozilla/5.0" & vbCrLf & _
  "Keep-Alive: 300" & vbCrLf & _
  "Connection: Keep -Alive" & vbCrLf & vbCrLf
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
  Dim Data As String
  Winsock1.GetData Data
  Data = Replace(Data, CrLf, vbCrLf)
  Text2 = Text2 & getcontent(Data)
End Sub

