VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Animation 
   Caption         =   "Smart Atm"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11925
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   562
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   10200
      Top             =   120
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   4560
      Picture         =   "Form2.frx":7ECB
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ATM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   5880
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   7440
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   11940
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
      uiMode          =   "none"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   21061
      _cy             =   13123
   End
End
Attribute VB_Name = "Animation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Dim s(5) As String


Private Sub Form_Load()
i = 0
s(0) = "E:\vbProject\video\ranvir.mp4"
s(1) = "E:\vbProject\video\Fedral.mp4"
s(2) = "E:\vbProject\video\Citi.mp4"
s(3) = "E:\vbProject\video\Fedral2.mp4"
End Sub

Private Sub Timer1_Timer()
WindowsMediaPlayer1.Controls.play
WindowsMediaPlayer1.Visible = True
'WindowsMediaPlayer1.Enabled = True

Timer1.Interval = 30000
WindowsMediaPlayer1.URL = s(i)
i = i + 1
If i > 3 Then
i = 0
End If
End Sub

Private Sub WindowsMediaPlayer1_Click(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
WindowsMediaPlayer1.Controls.stop
WindowsMediaPlayer1.Visible = False
AtmNumber.Timer1.Enabled = True
AtmNumber.Timer2.Enabled = True
AtmNumber.Timer3.Enabled = True
Timer1.Enabled = False
Timer1.Interval = 2000

Animation.Hide
AtmNumber.Show
i = 0
End Sub


