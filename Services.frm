VERSION 5.00
Begin VB.Form Services 
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
   Begin VB.Label Label4 
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
      Left            =   6000
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   4560
      Picture         =   "Services.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Postpaid Mobile Bill"
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
      Height          =   495
      Left            =   8280
      MouseIcon       =   "Services.frx":672D
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6600
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Trust Donation"
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
      Left            =   8280
      MouseIcon       =   "Services.frx":6BB7
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4800
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Pay"
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
      Height          =   495
      Left            =   8280
      MouseIcon       =   "Services.frx":7041
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   8415
      Left            =   0
      Picture         =   "Services.frx":74CB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "Services"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DonationBalance1, DonationBalance2 As Long
Dim s As String
Dim Conn As New ADODB.Connection
Dim Record As New ADODB.Recordset
Public flag As Integer

Private Sub Label1_Click()
flag = 1
ServicesPay.Label1.Caption = "Enter Your Bill No"
ServicesPay.Text1.MaxLength = 4
ServicesPay.Text1.Text = ""
ServicesPay.Show
Services.Hide
End Sub

Private Sub Label2_Click()
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\vbProject\Database\ATM.mdb;Persist Security Info=False"

s = "select Balance from Account where Acc_No= " & 112233405 & " or Acc_No=" & 112233406 & " "
Record.Open s, Conn, adOpenDynamic, adLockOptimistic
DonationBalance1 = Record.Fields(0)
Record.MoveNext
DonationBalance2 = Record.Fields(0)
Conn.Close
Donation.Show
Services.Hide
End Sub

Private Sub Label3_Click()
flag = 2
ServicesPay.Label1.Caption = "Enter Your Mobile no"
ServicesPay.Text1.MaxLength = 10
ServicesPay.Text1.Text = ""
ServicesPay.Show
Services.Hide
End Sub
