VERSION 5.00
Begin VB.Form BalanceRemaining 
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
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   10560
      Top             =   480
   End
   Begin VB.Label Label4 
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
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   4080
      Picture         =   "BalanceRemaining.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RS"
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
      Left            =   5160
      TabIndex        =   2
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   4920
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Remaining"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   3240
      TabIndex        =   0
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   8415
      Left            =   0
      Picture         =   "BalanceRemaining.frx":672D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "BalanceRemaining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As String
Dim Record As New ADODB.Recordset
Dim Conn As New ADODB.Connection

Private Sub Timer1_Timer()
Animation.Show
Animation.Timer1.Enabled = True
BalanceRemaining.Hide
Timer1.Enabled = False
End Sub

Public Function balance()
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\vbProject\Database\ATM.mdb;Persist Security Info=False"
s = "select Balance from Account where Card_No='" & AtmNumber.AccountCardNo & "' "
Record.Open s, Conn, adOpenDynamic, adLockOptimistic
Label2.Caption = Record.Fields(0)
Conn.Close
End Function
