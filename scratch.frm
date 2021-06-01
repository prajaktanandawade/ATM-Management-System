VERSION 5.00
Begin VB.Form scratch 
   BackColor       =   &H80000005&
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
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   11400
      Top             =   480
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   71
      Left            =   5880
      TabIndex        =   78
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   78
      Left            =   6480
      TabIndex        =   77
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   77
      Left            =   6480
      TabIndex        =   76
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   76
      Left            =   6720
      TabIndex        =   75
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   75
      Left            =   6720
      TabIndex        =   74
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   70
      Left            =   6120
      TabIndex        =   73
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   69
      Left            =   5880
      TabIndex        =   72
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   68
      Left            =   6240
      TabIndex        =   71
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   67
      Left            =   3720
      TabIndex        =   70
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   66
      Left            =   3960
      TabIndex        =   69
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   65
      Left            =   4200
      TabIndex        =   68
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   64
      Left            =   4440
      TabIndex        =   67
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   63
      Left            =   4680
      TabIndex        =   66
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   62
      Left            =   4920
      TabIndex        =   65
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   61
      Left            =   5160
      TabIndex        =   64
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   60
      Left            =   5400
      TabIndex        =   63
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   59
      Left            =   5640
      TabIndex        =   62
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   58
      Left            =   5160
      TabIndex        =   61
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   57
      Left            =   4920
      TabIndex        =   60
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   56
      Left            =   4680
      TabIndex        =   59
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   55
      Left            =   4440
      TabIndex        =   58
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   54
      Left            =   5400
      TabIndex        =   57
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   53
      Left            =   4200
      TabIndex        =   56
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   52
      Left            =   5640
      TabIndex        =   55
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Height          =   195
      Index           =   51
      Left            =   5880
      TabIndex        =   54
      Top             =   3960
      Width           =   45
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   50
      Left            =   6120
      TabIndex        =   53
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   49
      Left            =   3960
      TabIndex        =   52
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   48
      Left            =   3720
      TabIndex        =   51
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   47
      Left            =   3600
      TabIndex        =   50
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   46
      Left            =   3600
      TabIndex        =   49
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   45
      Left            =   6240
      TabIndex        =   48
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   44
      Left            =   6240
      TabIndex        =   47
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   43
      Left            =   5040
      TabIndex        =   46
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   42
      Left            =   5280
      TabIndex        =   45
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   41
      Left            =   3600
      TabIndex        =   44
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   40
      Left            =   4800
      TabIndex        =   43
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   39
      Left            =   4560
      TabIndex        =   42
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   38
      Left            =   4320
      TabIndex        =   41
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   37
      Left            =   5520
      TabIndex        =   40
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   36
      Left            =   4080
      TabIndex        =   39
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   35
      Left            =   3840
      TabIndex        =   38
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   34
      Left            =   5760
      TabIndex        =   37
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   33
      Left            =   6000
      TabIndex        =   36
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   32
      Left            =   6240
      TabIndex        =   35
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   31
      Left            =   6240
      TabIndex        =   34
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   30
      Left            =   3720
      TabIndex        =   33
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   29
      Left            =   3960
      TabIndex        =   32
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   28
      Left            =   4200
      TabIndex        =   31
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   27
      Left            =   4440
      TabIndex        =   30
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   26
      Left            =   4680
      TabIndex        =   29
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   25
      Left            =   4920
      TabIndex        =   28
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   24
      Left            =   5160
      TabIndex        =   27
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   23
      Left            =   5400
      TabIndex        =   26
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   22
      Left            =   5640
      TabIndex        =   25
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   21
      Left            =   5880
      TabIndex        =   24
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   20
      Left            =   6120
      TabIndex        =   23
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   19
      Left            =   3600
      TabIndex        =   22
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   18
      Left            =   6720
      TabIndex        =   21
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   17
      Left            =   6480
      TabIndex        =   20
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   16
      Left            =   6480
      TabIndex        =   19
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   15
      Left            =   6720
      TabIndex        =   18
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   14
      Left            =   6720
      TabIndex        =   17
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   13
      Left            =   6480
      TabIndex        =   16
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   12
      Left            =   6720
      TabIndex        =   15
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   11
      Left            =   6480
      TabIndex        =   14
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   10
      Left            =   3840
      TabIndex        =   13
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   9
      Left            =   4080
      TabIndex        =   12
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   8
      Left            =   4320
      TabIndex        =   11
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   7
      Left            =   4560
      TabIndex        =   10
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   9
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   5040
      TabIndex        =   8
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   6000
      TabIndex        =   7
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   5760
      TabIndex        =   6
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   5
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   4
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   3
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   2
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5175
      TabIndex        =   1
      Top             =   3960
      Width           =   105
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pleae Scratch Your Card"
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
      Left            =   3120
      TabIndex        =   0
      Top             =   1920
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   8415
      Left            =   0
      Picture         =   "scratch.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "scratch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conn As New ADODB.Connection
Dim Record As New ADODB.Recordset
Public Lowerbound, Upperbound As Integer
Dim s, s2 As String
Dim cashbackDate As Date
Public RandomNumber As Integer



Private Sub Label4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Label4(Index).Visible = False
End Sub

Public Function updateData()
cashbackDate = Format(Now, "dd/mm/yyyy")
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\vbProject\Database\ATM.mdb;Persist Security Info=False"
AtmNumber.AccountBalance = AtmNumber.AccountBalance + Val(Label2.Caption)
s = "update Account set Balance=" & AtmNumber.AccountBalance & " where Acc_No= " & AtmNumber.AccountNo & ""
s2 = "insert into Trans(T_Type,Acc_No,DOT,Narration,Amount) values(' & CashBack & '," & AtmNumber.AccountNo & ",'" & cashbackDate & "',' & cashback+ cr & '," & Val(Label2.Caption) & ")"
Record.Open s, Conn, adOpenDynamic, adLockOptimistic
Record.Open s2, Conn, adOpenDynamic, adLockOptimistic
Conn.Close
End Function
Public Function RandomMoney(ByVal x As Integer, ByVal y As Integer) As Integer
Randomize
RandomMoney = Int((y - x + 1) * Rnd + x)
End Function

Public Function RandomInteger()
    Timer1.Enabled = True
    Randomize
    RandomNumber = Int((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
    If (RandomNumber Mod 2 = 0) And (RandNumber Mod 5 = 0) Then
    Label2.Caption = RandomMoney(10, 100)
    Label3.Caption = "RS"
    Call updateData
    Else
    Label2.Caption = "Better Luck Next Time"
    Label3.Caption = ""
    End If
End Function
 
Private Sub Timer1_Timer()
PleaseWait.Timer1.Enabled = True
PleaseWait.Timer2.Enabled = True
PleaseWait.Timer3.Enabled = True
PleaseWait.Show
scratch.Hide
Timer1.Enabled = False
End Sub
