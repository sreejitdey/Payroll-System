VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Employee Login"
   ClientHeight    =   6465
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7215
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "FORGOT PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "REGISTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   7
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to our site"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "EMPLOYEE LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6735
   End
   Begin VB.Menu MnuSP 
      Caption         =   "Select Portal"
      Begin VB.Menu MnuEL 
         Caption         =   "Employee Login"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MnuAL 
         Caption         =   "Admin Login"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset

Private Sub Command1_Click()
If Text1.Text = "" And Text2.Text = "" Then
MsgBox ("Username and password required")
End
End If
rs.MoveFirst
Dim I As Integer
I = 0
While Not rs.EOF = True
If rs.Fields(0).Value = Text1.Text And rs.Fields(1).Value = Text2.Text Then
I = I + 1
Form4.Text1.Text = rs.Fields(2).Value
Form4.Text2.Text = rs.Fields(4).Value
Form4.Text3.Text = rs.Fields(5).Value
Form4.Text4.Text = rs.Fields(6).Value
Form4.Text5.Text = rs.Fields(7).Value
MsgBox ("Welcome to the portal")
Form4.Show
Form1.Hide
Form2.Hide
Form3.Hide
Form5.Hide
Form6.Hide
End If
rs.MoveNext
Wend
If I = 0 Then
MsgBox ("Invalid username or password")
Text1.Text = ""
Text2.Text = ""
End If
End Sub

Private Sub Command2_Click()
MsgBox ("Welcome to our registration portal")
Form3.Show
Form1.Hide
Form2.Hide
Form4.Hide
Form5.Hide
Form6.Hide
End Sub

Private Sub Command3_Click()
Form2.Show
Form1.Hide
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
End Sub

Private Sub Command4_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("G:\Visual Basic\PROJECT\Login_Database.mdb")
Set rs = db.OpenRecordset("select * from Table1")
End Sub

Private Sub MnuAL_Click()
Label1.Caption = "WELCOME TO THE ADMINISTRATOR"
Label2.Caption = "Please login to our site"
End Sub

Private Sub MnuEL_Click()
Label1.Caption = "EMPLOYEE LOGIN"
Label2.Caption = "Welcome to our site"
End Sub
