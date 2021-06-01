VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form PleaseWait 
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
   Begin VB.Timer Timer3 
      Interval        =   15000
      Left            =   4200
      Top             =   6840
   End
   Begin VB.Timer Timer2 
      Interval        =   4000
      Left            =   5040
      Top             =   6840
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   6000
      Top             =   6840
   End
   Begin VB.Image Image3 
      Height          =   1095
      Left            =   3960
      Picture         =   "PleaseWait.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label2 
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
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   6840
      Visible         =   0   'False
      Width           =   2655
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
      _cx             =   4683
      _cy             =   873
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait while Your Transaction is being processed"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   10455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   4170
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   1575
      Left            =   4800
      Picture         =   "PleaseWait.frx":672D
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   8415
      Left            =   0
      Picture         =   "PleaseWait.frx":11801
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "PleaseWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Integer

Private Sub Form_Load()
flag = 0
WindowsMediaPlayer1.URL = "C:\Users\Amit\Desktop\movies\vbProject\converted audio\speech.wav"
End Sub

Private Sub Timer1_Timer()
If flag = 0 Then
Image2.Picture = LoadPicture("E:\vbProject\images\Animation.bmp")
Shape1.Left = 336
Shape1.Top = 270
flag = 1
ElseIf flag = 1 Then
'Shape1.Left = 5760
Shape1.Top = 278
flag = 2

ElseIf flag = 2 Then
Shape1.Left = 352
Shape1.Top = 263
flag = 3
ElseIf flag = 3 Then
'Shape1.Left = 6480
Shape1.Top = 271
flag = 4

ElseIf flag = 4 Then
Shape1.Left = 370
Shape1.Top = 252
flag = 5

ElseIf flag = 5 Then
Shape1.Top = 260
flag = 6


ElseIf flag = 6 Then
Shape1.Left = 400
Shape1.Top = 236
flag = 7

ElseIf flag = 7 Then
Shape1.Top = 244
flag = 8

ElseIf flag = 8 Then
Shape1.Left = 424
Shape1.Top = 222
flag = 9

ElseIf flag = 9 Then
Shape1.Top = 230
flag = 10

ElseIf flag = 10 Then
Image2.Picture = LoadPicture("E:\vbProject\images\Animationcopy.bmp")
Shape1.Left = 422
Shape1.Top = 278
flag = 11

ElseIf flag = 11 Then
Shape1.Left = 400
Shape1.Top = 262
flag = 12

ElseIf flag = 12 Then
Shape1.Top = 270
flag = 13

ElseIf flag = 13 Then
Shape1.Left = 376
Shape1.Top = 250
flag = 14
ElseIf flag = 14 Then
Shape1.Top = 258
flag = 15

ElseIf flag = 15 Then
Shape1.Left = 352
Shape1.Top = 234
flag = 16
ElseIf flag = 16 Then
Shape1.Top = 242
flag = 17

ElseIf flag = 17 Then
Shape1.Left = 328
Shape1.Top = 222
flag = 18
ElseIf flag = 18 Then
Shape1.Top = 230
flag = 0
End If

End Sub

Private Sub Timer2_Timer()
WindowsMediaPlayer1.URL = "E:\vbProject\converted audio\speech.wav"
End Sub

Private Sub Timer3_Timer()
Timer2.Enabled = False
Timer1.Enabled = False
ThankYou.Timer1.Enabled = True
ThankYou.Label1.Caption = "Transaction Completed"
ThankYou.Show
PleaseWait.Hide
flag = 0
Timer3.Enabled = False
End Sub
