VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form AtmNumber 
   Caption         =   "Form3"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11925
   LinkTopic       =   "Form3"
   ScaleHeight     =   562
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9000
      Top             =   600
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9360
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3500
      Left            =   9720
      Top             =   600
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      MaxLength       =   16
      TabIndex        =   14
      Top             =   3240
      Width           =   4575
   End
   Begin VB.CommandButton Command13 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5520
      Picture         =   "Form3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5520
      Picture         =   "Form3.frx":3398
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5520
      Picture         =   "Form3.frx":6730
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      Picture         =   "Form3.frx":9AC8
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4560
      Picture         =   "Form3.frx":BD80
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3600
      Picture         =   "Form3.frx":E038
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2640
      Picture         =   "Form3.frx":102F0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4560
      Picture         =   "Form3.frx":125A8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3600
      Picture         =   "Form3.frx":14860
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2640
      Picture         =   "Form3.frx":16D18
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4560
      Picture         =   "Form3.frx":18FD0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3600
      Picture         =   "Form3.frx":1B288
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2640
      Picture         =   "Form3.frx":1D540
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ATM"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   5400
      TabIndex        =   20
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   3480
      Picture         =   "Form3.frx":1EF88
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label CardNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxx xxxx xxxx xxxx"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   19
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "sec"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   10680
      TabIndex        =   18
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   9720
      TabIndex        =   17
      Top             =   360
      Width           =   1335
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer2 
      Height          =   615
      Left            =   7680
      TabIndex        =   16
      Top             =   7320
      Visible         =   0   'False
      Width           =   3855
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6800
      _cy             =   1085
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   630
      Left            =   7560
      TabIndex        =   15
      Top             =   6600
      Visible         =   0   'False
      Width           =   3975
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   7011
      _cy             =   1111
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Enter Your 16 Digit Card Number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   13
      Top             =   1920
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   8415
      Left            =   0
      Picture         =   "Form3.frx":256B5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "AtmNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AccountBalance As Long
Public AccountCardNo As String
Public AccountNo As Long
Public AccountPin As Integer
Public AccountType As String
Public BankId As Integer
Dim Conn As New ADODB.Connection
Dim Record As New ADODB.Recordset
Dim s, status As String
Dim check As String
Dim flag, temp As Integer

Private Sub Command12_Click()
Call atmTimerDisable
Call restart
AtmNumber.Hide
End Sub

Private Sub Form_Load()
flag = 0
temp = 0
End Sub
Public Function restart()
Animation.Timer1.Enabled = True
Animation.Show
End Function
Public Function atmTimerDisable()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Label2.Caption = 60
Label2.ForeColor = vbWhite
Text1.Text = ""
End Function

Private Sub Command1_Click()
WindowsMediaPlayer2.URL = "E:\vbProject\converted audio\Windows.wav"
Text1.Text = Text1.Text + "1"
End Sub
Private Sub Command2_Click()
Text1.Text = Text1.Text + "2"
WindowsMediaPlayer2.URL = "E:\vbProject\converted audio\Windows.wav"
End Sub

Private Sub Command3_Click()
Text1.Text = Text1.Text + "3"
WindowsMediaPlayer2.URL = "E:\vbProject\converted audio\Windows.wav"
End Sub

Private Sub Command4_Click()
Text1.Text = Text1.Text + "4"
WindowsMediaPlayer2.URL = "E:\vbProject\converted audio\Windows.wav"
End Sub

Private Sub Command5_Click()
Text1.Text = Text1.Text + "5"
WindowsMediaPlayer2.URL = "E:\vbProject\converted audio\Windows.wav"
End Sub

Private Sub Command6_Click()
Text1.Text = Text1.Text + "6"
WindowsMediaPlayer2.URL = "E:\vbProject\converted audio\Windows.wav"
End Sub

Private Sub Command7_Click()
Text1.Text = Text1.Text + "7"
WindowsMediaPlayer2.URL = "E:\vbProject\converted audio\Windows.wav"
End Sub

Private Sub Command8_Click()
Text1.Text = Text1.Text + "8"
WindowsMediaPlayer2.URL = "E:\vbProject\converted audio\Windows.wav"
End Sub

Private Sub Command9_Click()
Text1.Text = Text1.Text + "9"
WindowsMediaPlayer2.URL = "E:\vbProject\converted audio\Windows.wav"
End Sub
Private Sub Command10_Click()
Text1.Text = Text1.Text + "0"
WindowsMediaPlayer2.URL = "E:\vbProject\converted audio\Windows.wav"
End Sub
Private Sub Command11_Click()
WindowsMediaPlayer2.URL = "E:\vbProject\converted audio\Windows.wav"
Text1.Text = ""
End Sub
Private Sub Command13_Click()
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\vbProject\Database\ATM.mdb;Persist Security Info=False"
s = "select * from Account where Card_No='" & Text1.Text & "' "
 Record.Open s, Conn, adOpenDynamic, adLockOptimistic
If Record.EOF Then
Call atmTimerDisable
Invalid.Label1.Caption = "Inavlid Card Number"
Invalid.Timer1.Enabled = True
Invalid.Show
AtmNumber.Hide
Else
AccountNo = Record.Fields(0)
AccountType = Record.Fields(1)
AccountCardNo = Record.Fields(2)
AccountBalance = Record.Fields(3)
BankId = Record.Fields(4)
status = Record.Fields(7)
If status = "active" Then
Call atmTimerDisable
AtmPin.Timer1.Enabled = True
AtmPin.Show
AtmPin.priintname
AtmNumber.Hide
Else
Call atmTimerDisable
Invalid.Label1.Caption = "You Card Has been Blocked"
Invalid.Timer1.Enabled = True
Invalid.Show
AtmNumber.Hide
End If
End If
Record.Close
Conn.Close
End Sub

Private Sub Timer1_Timer()
If flag = 0 Then
WindowsMediaPlayer1.URL = "E:\vbProject\converted audio\alice.wav"
flag = 1
Else
WindowsMediaPlayer1.URL = "E:\vbProject\converted audio\Hindi.wav"
flag = 0
Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
Label2.Caption = Label2.Caption - 1
If Label2.Caption <= 10 Then
Label2.ForeColor = vbRed
End If
If Label2.Caption <= 0 Then
Call atmTimerDisable
Animation.Show
Animation.Timer1.Enabled = True
AtmNumber.Hide
End If
End Sub

Private Sub Timer3_Timer()
If temp = 0 Then
CardNumber.ForeColor = vbWhite
temp = 1
ElseIf temp = 1 Then
CardNumber.ForeColor = vbBlack
temp = 0
End If
End Sub

