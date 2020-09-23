VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Francis B. Co's Main Window"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5505
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "MAIN"
      TabPicture(0)   =   "frmAdddeluser.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ADD OR DELETE A USER"
      TabPicture(1)   =   "frmAdddeluser.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboAccounttype"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdDEL"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtpassword2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtpassword1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdADD"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtUsername"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label5"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label4"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label3"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      Begin VB.ComboBox cboAccounttype 
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
         Left            =   -73200
         TabIndex        =   11
         Text            =   "USER"
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CommandButton cmdDEL 
         Caption         =   "DELETE"
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
         Left            =   -72240
         TabIndex        =   6
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtpassword2 
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
         IMEMode         =   3  'DISABLE
         Left            =   -73200
         PasswordChar    =   "*"
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtpassword1 
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
         IMEMode         =   3  'DISABLE
         Left            =   -73200
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CommandButton cmdADD 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73440
         TabIndex        =   3
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtUsername 
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
         Left            =   -73200
         TabIndex        =   2
         Text            =   "Type Here"
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "ACCOUNT TYPE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73200
         TabIndex        =   10
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "VERIFY PASSWORD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73200
         TabIndex        =   9
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   -73200
         TabIndex        =   8
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "USERNAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73200
         TabIndex        =   7
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   $"frmAdddeluser.frx":0038
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   600
         TabIndex        =   1
         Top             =   840
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdADD_Click()
Dim newaccount As String
Dim addpointer As String
Dim affirm As String

OpenDbase
newaccount = "[username] Like '*" & txtUsername & "*'"

With rs
.FindFirst newaccount
addpointer = !username

If (LCase(txtUsername) <> LCase(addpointer) And LCase(txtpassword1) = LCase(txtpassword2)) Then

affirm = MsgBox("Are you sure you want to add this account", vbYesNo)
  If (affirm = vbYes) Then

   .AddNew
   !username = LCase(txtUsername)
   !password = LCase(txtpassword1)
   !accounttype = LCase(cboAccounttype.Text)
   .Update
    MsgBox "New Account Added!", vbInformation
    .Close

  Else
  .Close
  Exit Sub
  End If

ElseIf (LCase(addpointer) = LCase(txtUsername)) Then
MsgBox "The username is already used, pick another!", vbOKOnly
.Close
Exit Sub

Else
MsgBox "Passwords do not match!", vbOKOnly
.Close
Exit Sub
End If
End With
End Sub

Private Sub cmdDEL_Click()
Dim oldaccount As String
Dim deletepointer As String
Dim affirm As String
OpenDbase
oldaccount = "[username] Like '*" & txtUsername & "*'"

With rs
.FindFirst oldaccount
deletepointer = !username

If (LCase(txtUsername) = LCase(deletepointer)) Then
affirm = MsgBox("Are you sure you want to delete this account?", vbYesNo)
 
 If (affirm = vbYes) Then
   .Delete
   MsgBox "Old account deleted!", vbInformation
   .Close
 Else
    Exit Sub
 End If
 
Else
MsgBox "Sorry I cannot find that account!", vbOKOnly
Exit Sub
End If
End With
End Sub

Private Sub Form_Load()
If (accounttype = "administrator") Then
cmdADD.Enabled = True
cmdDEL.Enabled = True
txtUsername.Enabled = True
txtpassword1.Enabled = True
txtpassword2.Enabled = True
cboAccounttype.Enabled = True
cboAccounttype.AddItem "USER"
cboAccounttype.AddItem "ADMINISTRATOR"
Else
MsgBox "This part cannot be accessed with a USER LEVEL account", vbOKOnly
cmdADD.Enabled = False
cmdDEL.Enabled = False
txtUsername.Enabled = False
txtpassword1.Enabled = False
txtpassword2.Enabled = False
cboAccounttype.Enabled = False
End If

End Sub
