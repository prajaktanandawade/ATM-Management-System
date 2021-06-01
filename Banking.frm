VERSION 5.00
Begin VB.Form Banking 
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
   Begin VB.Label Label6 
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
      Left            =   5160
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   3240
      Picture         =   "Banking.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pin Change"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      MouseIcon       =   "Banking.frx":672D
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "MiniStatement"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      MouseIcon       =   "Banking.frx":6BB7
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Enquiry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8760
      MouseIcon       =   "Banking.frx":7041
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Withdrawl"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9120
      MouseIcon       =   "Banking.frx":74CB
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fast Cash"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9000
      MouseIcon       =   "Banking.frx":7955
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   8415
      Left            =   0
      Picture         =   "Banking.frx":7DDF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "Banking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label1_Click()
FastCash.Show
Banking.Hide
End Sub

Private Sub Label2_Click()
Withdrawl.Show
Banking.Hide
End Sub

Private Sub Label3_Click()
BalanceRemaining.balance
BalanceRemaining.Timer1.Enabled = True
BalanceRemaining.Show
Banking.Hide
End Sub

Private Sub Label4_Click()
MiniStatement.mini
MiniStatement.printMini
MiniStatement.Show
Banking.Hide
End Sub

Private Sub Label5_Click()
PinChange.Label1.Caption = "Enter New Pin"
PinChange.Text1.Text = ""
PinChange.Show
Banking.Hide
End Sub
