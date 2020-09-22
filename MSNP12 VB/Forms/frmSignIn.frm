VERSION 5.00
Begin VB.Form frmSignIn 
   Caption         =   "Sign In - MSNP12 Visual Basic Example"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbStatus 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1335
      TabIndex        =   7
      Text            =   "Online"
      Top             =   1170
      Width           =   2790
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   300
      Left            =   5640
      TabIndex        =   5
      Top             =   1185
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sign In"
      Height          =   300
      Left            =   4230
      TabIndex        =   4
      Top             =   1185
      Width           =   1290
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   795
      Width           =   5565
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   465
      Width           =   5565
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright 2007 9InchWorM Software"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   150
      TabIndex        =   8
      Top             =   90
      Width           =   4560
   End
   Begin VB.Label Label1 
      Caption         =   "Initial Status:"
      Height          =   255
      Index           =   2
      Left            =   90
      TabIndex        =   6
      Top             =   1185
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Password:"
      Height          =   255
      Index           =   1
      Left            =   105
      TabIndex        =   1
      Top             =   825
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Email Adress:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1155
   End
End
Attribute VB_Name = "frmSignIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private stStatus As String
Private Const STATE_ONLINE = "NLN"
Private Const STATE_BUSY = "BSY"
Private Const STATE_BE_RIGHT_BACK = "BRB"
Private Const STATE_AWAY = "AWY"
Private Const STATE_ON_THE_PHONE = "PHN"
Private Const STATE_OUT_TO_LUNCH = "LUN"
Private Const STATE_HIDDEN = "HDN"
Private Const STATE_OFFLINE = "FLN"

Private Sub cmbStatus_Change()
If cmbStatus.Text = "Online" Then
stStatus = "STATE_ONLINE"
ElseIf cmbStatus.Text = "Busy" Then
stStatus = "STATE_BUSY"
ElseIf cmbStatus.Text = "Be Right Back" Then
stStatus = "STATE_BE_RIGHT_BACK"
ElseIf cmbStatus.Text = "Away" Then
stStatus = "STATE_AWAY"
ElseIf cmbStatus.Text = "On The Phone" Then
stStatus = "STATE_ON_THE_PHONE"
ElseIf cmbStatus.Text = "Out To Lunch" Then
stStatus = "STATE_OUT_TO_LUNCH"
ElseIf cmbStatus.Text = "Hidden" Then
stStatus = "STATE_HIDDEN"
ElseIf cmbStatus.Text = "Offline" Then
stStatus = "STATE_OFFLINE"
End If
End Sub

Private Sub Command1_Click()
Call frmMain.ConnectMSN(txtEmail.Text, txtPass.Text, STATE_AWAY)
Load frmMain
frmMain.Show
End Sub

Private Sub Form_Load()
stStatus = "STATE_ONLINE"
With cmbStatus
.AddItem "Online"
.AddItem "Busy"
.AddItem "Be Right Back"
.AddItem "Away"
.AddItem "On The Phone"
.AddItem "Out To Lunch"
.AddItem "Hidden"
.AddItem "Offline"
End With
End Sub
