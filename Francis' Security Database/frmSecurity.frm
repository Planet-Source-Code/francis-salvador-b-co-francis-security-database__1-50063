VERSION 5.00
Begin VB.Form frmSecurity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FRANCIS B.CO'S SECURITY DATABASE"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdquit 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ComboBox cbousername 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   840
      TabIndex        =   3
      Text            =   "Select User"
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "123"
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "USERNAME:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdlogin_Click()
Dim logger As String
Dim tempusername As String
Dim temppassword As String
Dim tempaccounttype As String

OpenDbase
logger = "[username] Like '*" & cbousername.Text & "*'"

With rs
.FindFirst logger
tempusername = !username
temppassword = !password
tempaccounttype = !accounttype
.Close
End With

If (LCase(tempusername) = LCase(cbousername.Text) And LCase(temppassword) = LCase(txtPassword)) Then

username = tempusername
password = temppassword
accounttype = tempaccounttype
frmmain.Show
Unload Me

Else
MsgBox "Sorry you cannot access, please try again!", vbOKOnly
Exit Sub

End If
End Sub

Private Sub cmdquit_Click()
MsgBox "THANK YOU FOR USING THE SAMPLE PROGRAM"
End
End Sub

Private Sub Form_Load()
OpenDbase
Do
cbousername.AddItem rs.Fields("username").Value
rs.MoveNext
Loop While Not rs.EOF
rs.Close
End Sub
