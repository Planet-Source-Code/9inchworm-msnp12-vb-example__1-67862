VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "MSNP12 Visual Basics Example / 9InchWorM Software 2007"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNC 
      Caption         =   "Set"
      Height          =   285
      Left            =   720
      TabIndex        =   11
      Top             =   6705
      Width           =   1500
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   60
      TabIndex        =   9
      Top             =   6375
      Width           =   3000
   End
   Begin VB.ListBox lstGroups 
      Height          =   840
      Left            =   3330
      TabIndex        =   8
      Top             =   6120
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.ListBox lstContacts 
      Height          =   5130
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   6390
   End
   Begin MSWinsockLib.Winsock sckSB 
      Left            =   6120
      Top             =   7215
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckNS 
      Left            =   5670
      Top             =   7230
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Change Nickname:"
      Height          =   225
      Left            =   75
      TabIndex        =   10
      Top             =   6135
      Width           =   1560
   End
   Begin VB.Label lblGroup 
      Caption         =   "Current Groups:"
      Height          =   225
      Left            =   2025
      TabIndex        =   7
      Top             =   6120
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label lblPMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "%PMessage%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1500
      TabIndex        =   5
      Top             =   570
      Visible         =   0   'False
      Width           =   4905
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Message:"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   75
      TabIndex        =   4
      Top             =   570
      Width           =   1395
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "%Status%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   630
      TabIndex        =   3
      Top             =   345
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   330
      Width           =   810
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   105
      TabIndex        =   1
      Top             =   90
      Width           =   810
   End
   Begin VB.Label lblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "%NickName%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   5970
   End
   Begin VB.Image imgMenuBG 
      Height          =   855
      Left            =   -195
      Picture         =   "frmMain.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6960
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "Status"
      Visible         =   0   'False
      Begin VB.Menu mnuOnline 
         Caption         =   "Online"
      End
      Begin VB.Menu mnuAway 
         Caption         =   "Away"
      End
      Begin VB.Menu mnuBRB 
         Caption         =   "Be Right Back"
      End
      Begin VB.Menu mnuOTL 
         Caption         =   "Out To Lunch"
      End
      Begin VB.Menu mnuOTP 
         Caption         =   "On The Phone"
      End
      Begin VB.Menu mnuHidden 
         Caption         =   "Appear Offline"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'Copyright 2007 9InchWorM Software, Brendan Thomas All Rights Reserved. Please leave credits, where credit is due
'Email: gollum_nz@hotmail.com
'Challenge code copyright 2005 Siebe "Inky" Tolsma
'Math functions copyright Troy Osborne
'===========================================================================

'==================
'To Do:
' - Add Display Picture Support
' - Add Recieving Personal Message Support
' - Add Conversations
' - Review code for optimized, and clean running performance.
'==================

'====================
'This is just a simple login, recieve contact list, set status and friendly name, and some information code snippet.
'This is not a fully functional client, and is only a snippet from my own client code
'This code was made in just under an hour.
'If you wish to add support email me: gollum_nz@hotmail.com
'====================
Option Explicit
Public ClientIDNo As String
Private Credentials As String
Private dTransID As Double
Public Username, Password As String
Private LocalStatus As String
Private LocalPassword As String
Private LocalName As String
Private Const STATE_ONLINE = "NLN"
Private Const STATE_BUSY = "BSY"
Private Const STATE_BE_RIGHT_BACK = "BRB"
Private Const STATE_AWAY = "AWY"
Private Const STATE_ON_THE_PHONE = "PHN"
Private Const STATE_OUT_TO_LUNCH = "LUN"
Private Const STATE_HIDDEN = "HDN"
Private Const STATE_OFFLINE = "FLN"

Private Sub cmdNC_Click()
Call ChangeFriendlyname(txtName.Text)
End Sub

Private Sub Form_Load()
'ClientIDNo = "1" 'Mobile Device
ClientIDNo = "805306412"
End Sub

Public Function TriID() As Double
If dTransID = 2 ^ 32 - 1 Then
 dTransID = 1
End If
TriID = dTransID
dTransID = dTransID + 1
End Function

Public Function ConnectMSN(ByVal sUsername As String, ByVal sPassword As String, ByVal sStatus As String)
LocalName = sUsername
LocalPassword = sPassword
LocalStatus = sStatus
lstContacts.Clear
sckNS.Close
sckNS.Connect "messenger.hotmail.com", 1863
frmWait.Show
frmSignIn.Hide
End Function

Public Function SendData(Data As String)
    Data = Replace(Data, "#", TriID)
    Call frmMain.sckNS.SendData(Data)
    If Mid(Data, Len(Data) - 1) = vbCrLf Then
        Debug.Print ("<<<: " & Mid(Data, 1, Len(Data) - 2))
    Else
        Debug.Print ("<<<: " & Data)
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lblStatus_Click()
PopupMenu mnuStatus
End Sub

Private Sub mnuAway_Click()
ChangeLocalStatus (STATE_AWAY)
End Sub

Private Sub mnuBRB_Click()
ChangeLocalStatus (STATE_BE_RIGHT_BACK)
End Sub

Private Sub mnuHidden_Click()
ChangeLocalStatus (STATE_HIDDEN)
End Sub

Private Sub mnuOnline_Click()
ChangeLocalStatus (STATE_ONLINE)
End Sub

Private Sub mnuOTL_Click()
ChangeLocalStatus (STATE_OUT_TO_LUNCH)
End Sub

Private Sub mnuOTP_Click()
ChangeLocalStatus (STATE_ON_THE_PHONE)
End Sub

Private Sub sckNS_Connect()
Call SendData("VER # MSNP12 MSNP11 MSNP10 CVR0" & vbCrLf)
End Sub

Private Sub sckNS_DataArrival(ByVal bytesTotal As Long)
Dim sCommand As String, sData As String, sbuffer As String, sParams() As String
Dim I As Long
    Do
        Call sckNS.PeekData(sbuffer, vbString, bytesTotal)
        If InStr(1, sbuffer, vbCrLf) = 0 Then
            Exit Sub
        End If
        If InStr(1, sbuffer, " ") Then
            sCommand = Mid$(sbuffer, 1, InStr(1, sbuffer, " ") - 1)
        Else
            sCommand = Mid$(sbuffer, 1, InStr(1, sbuffer, vbCrLf) - 1)
        End If
        If sCommand = "MSG" Or sCommand = "NOT" Then
            I = InStr(1, sbuffer, vbCrLf)
            sParams() = Split(Mid$(sbuffer, 1, I - 1), " ")
            If CLng(Len(Mid$(sbuffer, I + 2))) < CLng(sParams(3)) Then
                Exit Do
            End If
            sData = Mid(GetData(vbNullString, False, I + sParams(3) + 1), I + 2)
        Else
            sData = GetData(sbuffer)
            sParams() = Split(sData, " ")
            sData = vbNullString
        End If
        Call ProcessData(sParams, sData)
    Loop While sckNS.BytesReceived <> 0
End Sub

Private Sub ProcessData(sParams() As String, sPayload As String)
Dim sSubParams() As String

    Select Case sParams(0)
        Case "VER"
            Call SendData("CVR # 0x0409 winnt 5.1 i386 MSNMSGR 7.5.0311 msmsgs " & LocalName & vbCrLf)
            
        Case "CVR"
            Call SendData("USR # TWN I " & LocalName & vbCrLf)
         
        Case "XFR"
            If sParams(2) = "NS" Then
                sSubParams() = Split(sParams(3), ":")
                Call sckNS.Close
                Call sckNS.Connect(sSubParams(0), sSubParams(1))
            End If
            
        Case "USR"
            If sParams(2) = "TWN" Then
                 Credentials = pKey(sParams(4), LocalName, LocalPassword)
                 If Credentials <> "" Then Call SendData("USR # TWN S " & Credentials & vbCrLf)
                 
            ElseIf sParams(2) = "OK" Then
                Call SendData("SYN # 0 0" & vbCrLf)
            End If
                 
        Case "SYN":
            Call SendData("CHG # " & LocalStatus & " " & ClientIDNo & vbCrLf)
            Call SendData("PNG" & vbCrLf)
            frmWait.Hide
            Call SetPersonalMessage("Testing")
            'Call SetCurrentMedia("Testing The Music", "Album 65", "9InchWorM Software")
            
        Case "PRP"
            Select Case sParams(1)
                Case "MFN"
                lblName.Visible = True
                lblName.Caption = URLDecode(sParams(2))
                Case "PHH"
                'Home Phone
                Case "PHW"
                'Work Phone
                Case "PHM"
                'Mobile Phone
                Case "MBE"
                'User has Mobile True/False
            End Select
        If sParams(2) = "MFN" Then
        lblName.Caption = URLDecode(sParams(3))
        txtName.Text = "Success!!"
        End If
            
        Case "CHL"
        Dim strInput As String
        strInput = CreateQRY(sParams(2))
        Call SendData("QRY # PROD0090YUAUV{2B 32" & vbCrLf & strInput)
        
        Case "LSG"
        ' Add Your Contacts Groups Code Here
        lstGroups.Visible = True
        lblGroup.Visible = True
        lstGroups.AddItem URLDecode(sParams(1))
            
        Case "UBX"
        ' ---
        
        Case "UUX"
        lblPMessage.Visible = True
        'Message Length
        lblPMessage.Caption = sParams(2)
                
        Case "LST"
        Dim strC As String
        strC = Replace(sParams(2), "F=", "")
        lstContacts.AddItem MSNDecode(strC)
        
        Case "MSG"
        '-- MSG Payload
        
        Case "NLN"
        If sParams(1) = "NLN" Then
        lstContacts.AddItem "Online: " & sParams(2)
        MsgBox sParams(2) & " has just signed in"
        Else
        ' Whatever
        End If
        
        Case "ILN"
        Dim strData As String
        strData = sParams(3)
        'lstContacts.AddItem "Online: " & (strData)
        
        Case "FLN"
        ' User has Signed off
        MsgBox sParams(1) & " is now offline."

        Case "CHG"
        lblStatus.Visible = True
            Select Case sParams(2)
                Case "HDN"
                    lblStatus.Caption = "Appear Offline"
                Case "AWY"
                    lblStatus.Caption = "Away"
                Case "BRB"
                    lblStatus.Caption = "Be Right Back"
                Case "BSY"
                    lblStatus.Caption = "Busy"
                Case "NLN"
                    lblStatus.Caption = "Online"
                Case "PHN"
                    lblStatus.Caption = "On The Phone"
                Case "LUN"
                    lblStatus.Caption = "Out To Lunch"
            End Select

        Case "RNG"
        MsgBox MSNDecode(sParams(6)) & " " & MSNDecode(sParams(5)) & " has opened up a chat window with you."
            ' Contact is trying to call you into a conversation
            
        Case "OUT"
            If sParams(1) = "OTH" Then
            MsgBox "You have been signed out because you signed in from another location."
            End If
    End Select
End Sub

Private Function GetData(sbuffer As String, Optional bTrim As Boolean = True, Optional lLength As Long = 0) As String
Dim sData As String
Dim I As Long

    If lLength = 0 Then
        I = InStr(1, sbuffer, vbCrLf, vbTextCompare)
        Call sckNS.GetData(sData, vbString, I + 1)
    Else
        Call sckNS.GetData(sData, vbString, lLength)
    End If
    If bTrim = True Then
        GetData = Mid(sData, 1, Len(sData) - 2)
    Else
        GetData = sData
    End If
    Debug.Print ">>>: " & GetData
End Function
    
    'Set Personal message for local user
Private Function ChangeLocalStatus(sStatus As String)
        If sStatus = STATE_ONLINE Then
            Call SendData("CHG " & TriID & " " & STATE_ONLINE & " 0" & vbCrLf)
        ElseIf sStatus = STATE_BUSY Then
            Call SendData("CHG " & TriID & " " & STATE_BUSY & " 0" & vbCrLf)
        ElseIf sStatus = STATE_BE_RIGHT_BACK Then
            Call SendData("CHG " & TriID & " " & STATE_BE_RIGHT_BACK & " 0" & vbCrLf)
        ElseIf sStatus = STATE_AWAY Then
            Call SendData("CHG " & TriID & " " & STATE_AWAY & " 0" & vbCrLf)
        ElseIf sStatus = STATE_HIDDEN Then
            Call SendData("CHG " & TriID & " " & STATE_HIDDEN & " 0" & vbCrLf)
        ElseIf sStatus = STATE_OUT_TO_LUNCH Then
            Call SendData("CHG " & TriID & " " & STATE_OUT_TO_LUNCH & " 0" & vbCrLf)
        ElseIf sStatus = STATE_ON_THE_PHONE Then
            Call SendData("CHG " & TriID & " " & STATE_ON_THE_PHONE & " 0" & vbCrLf)
        End If
End Function
