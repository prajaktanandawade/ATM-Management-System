VERSION 5.00
Begin VB.Form Homepage 
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":0000
   MousePointer    =   1  'Arrow
   ScaleHeight     =   562
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label8 
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
      Left            =   5880
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   3600
      Picture         =   "Form1.frx":048A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select your Transaction"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   6
      Top             =   1680
      Width           =   6615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " TRANSFER"
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
      Height          =   495
      Left            =   8760
      MouseIcon       =   "Form1.frx":6BB7
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FAST CASH"
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
      Height          =   495
      Left            =   0
      MouseIcon       =   "Form1.frx":7041
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   6720
      Width           =   3735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MINISTATEMENT"
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
      Height          =   615
      Left            =   0
      MouseIcon       =   "Form1.frx":74CB
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   4920
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "BALANCE ENQUIRY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   9000
      MouseIcon       =   "Form1.frx":7955
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BANKING"
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
      Height          =   495
      Left            =   8400
      MouseIcon       =   "Form1.frx":7DDF
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SERVICES"
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
      Height          =   615
      Left            =   0
      MouseIcon       =   "Form1.frx":8269
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   8535
      Left            =   0
      Picture         =   "Form1.frx":86F3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "Homepage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Balance1, Balance2 As Long
Dim s As String
Dim Record As New ADODB.Recordset
Dim Conn As New ADODB.Connection

Private Sub Label1_Click()
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\vbProject\Database\ATM.mdb;Persist Security Info=False"

s = "select Balance from Account where Acc_No=" & 112233403 & " or Acc_No=" & 112233404 & " "
Record.Open s, Conn, adOpenDynamic, adLockOptimistic
Balance1 = Record.Fields(0)
Record.MoveNext
Balance2 = Record.Fields(0)
Services.Show
Homepage.Hide
Conn.Close
End Sub

Private Sub Label2_Click()
Banking.Show
Homepage.Hide
End Sub

Private Sub Label3_Click()
BalanceRemaining.balance
BalanceRemaining.Timer1.Enabled = True
BalanceRemaining.Show
Homepage.Hide
End Sub

Private Sub Label4_Click()
MiniStatement.Show
MiniStatement.mini
MiniStatement.printMini
Homepage.Hide
End Sub

Private Sub Label5_Click()
FastCash.Show
Homepage.Hide
End Sub

Private Sub Label6_Click()
Transfer.Show
Homepage.Hide
End Sub
