VERSION 5.00
Begin VB.Form ServicesAmount 
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
      Left            =   5760
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   4440
      Picture         =   "ServicesAmount.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   8640
      TabIndex        =   4
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label4 
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
      Height          =   615
      Left            =   7200
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
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
      Height          =   375
      Left            =   9840
      MouseIcon       =   "ServicesAmount.frx":672D
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PAY"
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
      Height          =   375
      Left            =   9720
      MouseIcon       =   "ServicesAmount.frx":6BB7
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   8415
      Left            =   0
      Picture         =   "ServicesAmount.frx":7041
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "ServicesAmount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s1, s2, s3, s4, s5, narr1, narr2, TransType1, TransType2, stat As String
Dim BillDate As Date
Dim Record As New ADODB.Recordset
Dim Conn As New ADODB.Connection

Private Sub Command1_Click()
End Sub

Private Sub Form_Load()
stat = "completed"
BillDate = Format(Now, "dd/mm/yyyy")
TransType1 = "Electricity Bill Pay"
TransType2 = "Mobile Bill Pay"
narr1 = "Bill Payment - debited"
narr2 = "Bill payement +credited "
End Sub

Private Sub Label2_Click()
If AtmNumber.AccountBalance > Label1.Caption Then

Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\vbProject\Database\ATM.mdb;Persist Security Info=False"

AtmNumber.AccountBalance = AtmNumber.AccountBalance - Label1.Caption
s1 = "Update Account set Balance=" & AtmNumber.AccountBalance & " where Acc_No=" & AtmNumber.AccountNo & "  "

If ServicesPay.flag = 1 Then
stat = "Completed"
Homepage.Balance1 = ServicesPay.BillAmount + Homepage.Balance1
s2 = "Update Account set Balance=" & Homepage.Balance1 & " where Acc_No=" & 112233403 & " "
s3 = "insert into Trans(T_Type,Acc_No,DOT,Narration,Amount) values('" & TransType1 & "'," & AtmNumber.AccountNo & ",'" & BillDate & "','" & narr1 & "'," & Label1.Caption & ")"
s4 = "insert into Trans(T_Type,Acc_No,DOT,Narration,Amount) values('" & TransType1 & "'," & 112233403 & ",'" & BillDate & "','" & narr2 & "'," & Label1.Caption & ")"
s5 = "update BillPay set Status='" & stat & "' where Consumer_No=" & Val(ServicesPay.Text1.Text) & " "
Record.Open s1, Conn, adOpenDynamic, adLockOptimistic
Record.Open s2, Conn, adOpenDynamic, adLockOptimistic
Record.Open s3, Conn, adOpenDynamic, adLockOptimistic
Record.Open s4, Conn, adOpenDynamic, adLockOptimistic
Record.Open s5, Conn, adOpenDynamic, adLockOptimistic
ElseIf ServicesPay.flag = 2 Then
stat = "completed"
Homepage.Balance2 = ServicesPay.MobileBill + Homepage.Balance2
s2 = "Update Account set Balance=" & Homepage.Balance2 & " where Acc_No=" & 112233404 & " "
s3 = "insert into Trans(T_Type,Acc_No,DOT,Narration,Amount) values('" & TransType2 & "'," & AtmNumber.AccountNo & ",'" & BillDate & "','" & narr1 & "'," & Label1.Caption & ")"
s4 = "insert into Trans(T_Type,Acc_No,DOT,Narration,Amount) values('" & TransType2 & "'," & 112233404 & ",'" & BillDate & "','" & narr2 & "'," & Label1.Caption & ")"
s5 = "update MobilePay set Status='" & stat & "' where Mobile_No='" & ServicesPay.Text1.Text & "' "
Record.Open s1, Conn, adOpenDynamic, adLockOptimistic
Record.Open s2, Conn, adOpenDynamic, adLockOptimistic
Record.Open s3, Conn, adOpenDynamic, adLockOptimistic
Record.Open s4, Conn, adOpenDynamic, adLockOptimistic
Record.Open s5, Conn, adOpenDynamic, adLockOptimistic
End If
End If

Conn.Close
scratch.Upperbound = 5000
scratch.Lowerbound = 100
scratch.RandomInteger
scratch.Show
ServicesAmount.Hide
End Sub

Private Sub Label3_Click()
Animation.Show
Animation.Timer1.Enabled = True
ServicesAmount.Hide
End Sub
