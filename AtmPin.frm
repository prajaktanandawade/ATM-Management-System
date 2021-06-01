VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form AtmPin 
   Caption         =   "Form4"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11925
   LinkTopic       =   "Form4"
   ScaleHeight     =   562
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   11400
      Top             =   1200
   End
   Begin VB.CommandButton Command13 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7200
      Picture         =   "AtmPin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7200
      Picture         =   "AtmPin.frx":3398
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7200
      Picture         =   "AtmPin.frx":6730
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
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
      Left            =   5280
      Picture         =   "AtmPin.frx":9AC8
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton Command9 
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
      Left            =   6240
      Picture         =   "AtmPin.frx":BD80
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton Command8 
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
      Left            =   5280
      Picture         =   "AtmPin.frx":E038
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton Command7 
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
      Left            =   4320
      Picture         =   "AtmPin.frx":102F0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton Command6 
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
      Left            =   6240
      Picture         =   "AtmPin.frx":125A8
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Command5 
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
      Left            =   5280
      Picture         =   "AtmPin.frx":14860
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Command4 
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
      Left            =   4320
      Picture         =   "AtmPin.frx":16D18
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Command3 
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
      Left            =   6240
      Picture         =   "AtmPin.frx":18FD0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
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
      Left            =   5280
      Picture         =   "AtmPin.frx":1B288
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
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
      Left            =   4320
      Picture         =   "AtmPin.frx":1D540
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   4200
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2280
      Width           =   4575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
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
      Height          =   975
      Left            =   5400
      TabIndex        =   22
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   3360
      Picture         =   "AtmPin.frx":1EF88
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   8520
      TabIndex        =   21
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   8640
      TabIndex        =   20
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Attempts Left"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   375
      Left            =   9600
      TabIndex        =   19
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9360
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Wrong Pin "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   9240
      TabIndex        =   17
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer2 
      Height          =   375
      Left            =   8880
      TabIndex        =   16
      Top             =   7440
      Visible         =   0   'False
      Width           =   3015
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
      _cx             =   5318
      _cy             =   661
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   8760
      TabIndex        =   15
      Top             =   7920
      Visible         =   0   'False
      Width           =   3135
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
      _cx             =   5530
      _cy             =   873
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Enter Your 4 Digit Pin Number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   1680
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   8415
      Left            =   0
      Picture         =   "AtmPin.frx":256B5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "AtmPin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim Record2 As New ADODB.Recordset
Dim Record As New ADODB.Recordset
Dim Record3 As New ADODB.Recordset
Dim Conn As New ADODB.Connection
Dim s, s2, s3 As String
Dim flag As Integer

Private Sub Command1_Click()
WindowsMediaPlayer2.URL = "E:\vbProject\converted audio\Windows.wav"
Text1.Text = Text1.Text + "1"
End Sub

Private Sub Command12_Click()
Timer1.Enabled = False
AtmNumber.restart
AtmPin.Hide
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label3.Caption = 3
Text1.Text = ""
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
s = "select Pin from Account where Pin= " & Val(Text1.Text) & " And Card_No='" & AtmNumber.AccountCardNo & "' "
Record.Open s, Conn, adOpenDynamic, adLockOptimistic
If Record.EOF Then
Text1.Text = ""
Label3.Caption = Label3.Caption - 1
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
If Label3.Caption < 1 Then
Label3.Caption = 3
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Timer1.Enabled = False
Blocked.Visible = True
Blocked.Timer1.Enabled = True
AtmPin.Visible = False
End If
Else
Timer1.Enabled = False
Homepage.Show
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label3.Caption = 3
AtmPin.Visible = False
Text1.Text = ""
End If
Record.Close
Conn.Close
End Sub


Private Sub Timer1_Timer()
If flag = 0 Then
WindowsMediaPlayer1.URL = "E:\vbProject\converted audio\Alice2.wav"
flag = 1
Else
WindowsMediaPlayer1.URL = "E:\vbProject\converted audio\Atm hindi.wav"
Timer1.Enabled = False
flag = 0
End If
End Sub
Public Function priintname()
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\vbProject\Database\ATM.mdb;Persist Security Info=False"
s2 = "select c.Cust_Name from Customer c,Account a where c.Cust_Id=a.Cust_Id and a.Card_No='" & AtmNumber.AccountCardNo & "' "
Record2.Open s2, Conn, adOpenDynamic, adLockOptimistic
Label6.Caption = Record2.Fields(0)
Conn.Close
End Function

