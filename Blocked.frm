VERSION 5.00
Begin VB.Form Blocked 
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   10440
      Top             =   360
   End
   Begin VB.Label Label2 
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
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   3720
      Picture         =   "Blocked.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your Card Has Been Blocked. Please Contact Your Bank"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   5280
      Width           =   10335
   End
   Begin VB.Image Image1 
      Height          =   8415
      Left            =   0
      Picture         =   "Blocked.frx":672D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "Blocked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s, block As String
Dim Conn As New ADODB.Connection
Dim Record As New ADODB.Recordset

Private Sub Form_Load()
block = "blocked"
End Sub

Private Sub Timer1_Timer()
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\vbProject\Database\ATM.mdb;Persist Security Info=False"

s = "update Account set status='" & block & "' where Acc_No=" & AtmNumber.AccountNo & " "
Record.Open s, Conn, adOpenDynamic, adLockOptimistic
Animation.Show
Animation.Timer1.Enabled = True
Blocked.Hide
Conn.Close
Timer1.Enabled = False
End Sub
