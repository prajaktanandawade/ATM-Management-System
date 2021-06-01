VERSION 5.00
Begin VB.Form Donation 
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
      Height          =   1095
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   3600
      Picture         =   "Donation.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Smile Foundation"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      MouseIcon       =   "Donation.frx":672D
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   6720
      Width           =   3255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Alpha Foundation"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   240
      MouseIcon       =   "Donation.frx":6BB7
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   8415
      Left            =   0
      Picture         =   "Donation.frx":7041
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "Donation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public flag As Integer
Private Sub Label1_Click()
flag = 1
DonationPay.Show
Donation.Hide
End Sub

Private Sub Label2_Click()
flag = 2
DonationPay.Show
Donation.Hide
End Sub
