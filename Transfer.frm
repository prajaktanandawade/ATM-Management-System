VERSION 5.00
Begin VB.Form Transfer 
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
   Begin VB.Label Label3 
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
      Height          =   975
      Left            =   6240
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   4680
      Picture         =   "Transfer.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "acount transfer"
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
      Height          =   495
      Left            =   8400
      MouseIcon       =   "Transfer.frx":672D
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "card to card transfer"
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
      Height          =   615
      Left            =   8280
      MouseIcon       =   "Transfer.frx":6BB7
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   8415
      Left            =   0
      Picture         =   "Transfer.frx":7041
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "Transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public flag As Integer
Private Sub Label1_Click()
flag = 0
TransferCheck.Show
TransferCheck.Label1.Caption = "Enter 16 digit card no"
Transfer.Hide
End Sub

Private Sub Label2_Click()
flag = 1
TransferCheck.Show
TransferCheck.Label1.Caption = "Enter 9 digit Account no"
Transfer.Hide
End Sub
