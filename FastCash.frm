VERSION 5.00
Begin VB.Form FastCash 
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
   Begin VB.Label Label8 
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
      Height          =   855
      Left            =   5520
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   3960
      Picture         =   "FastCash.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "10000  Rs"
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
      Left            =   8880
      MouseIcon       =   "FastCash.frx":672D
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "5000  Rs"
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
      Left            =   8760
      MouseIcon       =   "FastCash.frx":6BB7
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "3000  Rs"
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
      Left            =   8640
      MouseIcon       =   "FastCash.frx":7041
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "1000  Rs"
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
      Left            =   480
      MouseIcon       =   "FastCash.frx":74CB
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "200  Rs"
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
      Left            =   360
      MouseIcon       =   "FastCash.frx":7955
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "100  Rs"
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
      Left            =   360
      MouseIcon       =   "FastCash.frx":7DDF
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select Your Amount"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   1680
      Width           =   6135
   End
   Begin VB.Image Image1 
      Height          =   8535
      Left            =   0
      Picture         =   "FastCash.frx":8269
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12375
   End
End
Attribute VB_Name = "FastCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChangeBalance As Long
Dim s, s2, narr, Trans_Type As String
Dim d As Date
Dim Record As New ADODB.Recordset
Dim Conn As New ADODB.Connection
Function fastCashAccount(money As Integer)
If money < AtmNumber.AccountBalance Then
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\vbProject\Database\ATM.mdb;Persist Security Info=False"

ChangeBalance = AtmNumber.AccountBalance - money
s = "update Account set Balance=" & ChangeBalance & " where Acc_no=" & AtmNumber.AccountNo & " "
s2 = "insert into Trans(T_Type,Acc_No,DOT,Narration,Amount) values('" & Trans_Type & "'," & AtmNumber.AccountNo & ", '" & d & "','" & narr & "'," & money & ") "
Record.Open s2, Conn, adOpenDynamic, adLockOptimistic
Record.Open s, Conn, adOpenDynamic, adLockOptimistic
PleaseWait.Timer1.Enabled = True
PleaseWait.Timer2.Enabled = True
PleaseWait.Timer3.Enabled = True
PleaseWait.Show
FastCash.Hide
Conn.Close
Else
Invalid.Show
Invalid.Timer1.Enabled = True
Invalid.Label1.Caption = "Insufficient Amount"
FastCash.Hide
End If
End Function

Private Sub Form_Load()
narr = "fast cash withdrwal"
d = Format(Now, "mm/dd/yyyy")
Trans_Type = "Fast Cash"
End Sub

Private Sub Label2_Click()
Call fastCashAccount(100)
End Sub

Private Sub Label3_Click()
Call fastCashAccount(200)
End Sub

Private Sub Label4_Click()
Call fastCashAccount(1000)
End Sub

Private Sub Label5_Click()
Call fastCashAccount(3000)
End Sub

Private Sub Label6_Click()
Call fastCashAccount(5000)
End Sub

Private Sub Label7_Click()
Call fastCashAccount(10000)
End Sub
