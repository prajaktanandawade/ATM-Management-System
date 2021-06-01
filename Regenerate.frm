VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Regenerate 
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   562
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   795
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command13 
      Caption         =   "enter"
      Height          =   615
      Left            =   7800
      TabIndex        =   13
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command12 
      Caption         =   "clear"
      Height          =   615
      Left            =   7800
      TabIndex        =   12
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   7800
      TabIndex        =   11
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   10
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   9
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   8
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   7
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   6
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   5
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   4
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   3
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   2
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   1
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   2160
      Width           =   3975
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   255
      Left            =   8880
      TabIndex        =   14
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
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
      volume          =   50
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
      _cx             =   1720
      _cy             =   450
   End
   Begin VB.Image Image1 
      Height          =   8415
      Left            =   0
      Picture         =   "Regenerate.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "Regenerate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
WindowsMediaPlayer1.URL = "E:\vbProject\converted audio\Windows.wav"
Text1.Text = Text1.Text + Command1.Caption
End Sub

Private Sub Command2_Click()
Text1.Text = Text1.Text + Command2.Caption
WindowsMediaPlayer1.URL = "E:\vbProject\converted audio\Windows.wav"
End Sub

Private Sub Command3_Click()
Text1.Text = Text1.Text + Command3.Caption
WindowsMediaPlayer1.URL = "E:\vbProject\converted audio\Windows.wav"
End Sub

Private Sub Command4_Click()
Text1.Text = Text1.Text + Command4.Caption
WindowsMediaPlayer1.URL = "E:\vbProject\converted audio\Windows.wav"

End Sub

Private Sub Command5_Click()
Text1.Text = Text1.Text + Command5.Caption
WindowsMediaPlayer1.URL = "E:\vbProject\converted audio\Windows.wav"

End Sub

Private Sub Command6_Click()
Text1.Text = Text1.Text + Command6.Caption
WindowsMediaPlayer1.URL = "E:\vbProject\converted audio\Windows.wav"

End Sub

Private Sub Command7_Click()
Text1.Text = Text1.Text + Command7.Caption
WindowsMediaPlayer1.URL = "E:\vbProject\converted audio\Windows.wav"

End Sub

Private Sub Command8_Click()
Text1.Text = Text1.Text + Command8.Caption
WindowsMediaPlayer1.URL = "E:\vbProject\converted audio\Windows.wav"

End Sub

Private Sub Command9_Click()
Text1.Text = Text1.Text + Command9.Caption
WindowsMediaPlayer1.URL = "E:\vbProject\converted audio\Windows.wav"

End Sub
Private Sub Command10_Click()
Text1.Text = Text1.Text + Command10.Caption
WindowsMediaPlayer1.URL = "E:\vbProject\converted audio\Windows.wav"

End Sub
Private Sub Command11_Click()
WindowsMediaPlayer1.URL = "E:\vbProject\converted audio\Windows.wav"
Text1.Text = ""
End Sub

