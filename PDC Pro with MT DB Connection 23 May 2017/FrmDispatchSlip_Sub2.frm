VERSION 5.00
Begin VB.Form FrmDispatchSlip_Sub2 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dispatch Slip Cancellation Entry"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3000
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtSlipDate 
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
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox TxtUserName 
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
      Height          =   315
      Left            =   240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtSlipNo 
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
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Accept"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtFlag 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "ADD"
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3000
      TabIndex        =   10
      Top             =   1080
      Width           =   825
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Entry By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dispatch Slip Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   210
      TabIndex        =   6
      Top             =   240
      Width           =   1875
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dispatch Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   1245
   End
End
Attribute VB_Name = "FrmDispatchSlip_Sub2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
On Error Resume Next
Unload Me
End Sub


Private Sub cmdSave_Click()
On Error Resume Next

If CheckValid = False Then Exit Sub

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

'Validate login details
str1 = "Select * from ASPDC_UserMaster where PDCUserName ='" & Me.TxtUserName.Text & "' and PDCPassword ='" & Me.txtPass1.Text & "' and ActiveStatus = 1"
rs1.Open str1, cn1, adOpenDynamic, adLockReadOnly
If (rs1.BOF And rs1.EOF) Then
    MsgBox "Invalid Password", vbCritical + vbOKOnly
    txtPass1.SetFocus
    rs1.Close
    cn1.Close
    Exit Sub
End If
rs1.Close

Dim str As String
    
str = "Select * from ASPDC_DispatchSlip where DispatchSlipCode ='" & Trim(UCase(Me.txtSlipNo.Text)) & "'"
rs1.Open str, cn1, adOpenDynamic, adLockReadOnly
If rs1.BOF And rs1.EOF Then
    AddSlip
Else
    MsgBox "The Dispatch Slip Number already exists.", vbInformation + vbOKOnly, "Error"
    txtSlipNo.SetFocus
    rs1.Close
    Exit Sub
End If
rs1.Close
cn1.Close

End Sub

Private Sub AddSlip()
On Error Resume Next
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset


rs1.Open "Select * from ASPDC_DispatchSlip where DispatchSlipCode ='" & txtSlipNo.Text & "'", cn1, adOpenDynamic, adLockPessimistic
rs1.AddNew
rs1.Fields("DispatchSlipCode").Value = Trim(txtSlipNo.Text)
rs1.Fields("MISInstituteCode").Value = Left(Trim(txtSlipNo.Text), 3)
rs1.Fields("LICCode").Value = Mid(Trim(txtSlipNo.Text), 4, 2)
rs1.Fields("DispatchDate").Value = DateValue(Left(txtSlipDate.Text, 2) & "-" & MonthName(Mid(txtSlipDate.Text, 3, 2)) & "-" & Right(txtSlipDate.Text, 4))
rs1.Fields("ChequeCnt").Value = 0
rs1.Fields("ChequeValue").Value = 0
rs1.Fields("SlipStatus").Value = 2     'Cancelled
rs1.Fields("SlipEntryUserName").Value = TxtUserName.Text
rs1.Fields("SlipEntryDate").Value = Date

rs1.Fields("ChqEntryFlag").Value = 0
rs1.Fields("ChqEntryUserName").Value = ""
rs1.Fields("ChqEntryDate").Value = ""

rs1.Fields("AuthEntryFlag").Value = 0
rs1.Fields("AuthEntryUserName").Value = ""
rs1.Fields("AuthEntryDate").Value = ""

rs1.Fields("AutoMapChqCnt").Value = 0
rs1.Fields("ManualMapChqCnt").Value = 0
rs1.Fields("OpenChqCnt").Value = 0

rs1.Fields("CompleteEntryFlag").Value = 1   'Entry is complete as no more cheque entries

rs1.Update
rs1.Close

MsgBox "Dispatch Slip Cancellation Entry successfully done.", vbOKOnly + vbInformation
Unload Me
End Sub



Private Sub Form_Load()
On Error Resume Next
txtFlag.Text = "ADD"
TxtUserName.Text = ModInit.PDCUserName
End Sub


Private Function CheckValid() As Boolean
On Error Resume Next
If Trim(Me.txtSlipNo.Text) = "" Then
    MsgBox "Enter Dispatch Slip Number.", vbInformation + vbOKOnly
    txtSlipNo.SetFocus
    CheckValid = False
    Exit Function
End If

If Len(Trim(txtSlipNo.Text)) <> 13 Then
    MsgBox "Invalid Dispatch Slip Number.", vbCritical + vbOKOnly
    txtSlipNo.SetFocus
    CheckValid = False
    Exit Function
End If

CheckValid = True
End Function




Private Sub txtChequeCnt_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    txtChequeEntry.SetFocus
    KeyAscii = 0
End If
End Sub

Private Sub txtPass1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cmdSave_Click
    KeyAscii = 0
End If
End Sub

Private Sub txtSlipDate_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    Me.txtPass1.SetFocus
    KeyAscii = 0
End If
End Sub

Private Sub txtSlipNo_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    If Len(txtSlipNo.Text) = 13 Then
        txtSlipDate.SetFocus
    Else
        MsgBox "Invalid Slip No.", vbCritical + vbOKOnly
    End If
    KeyAscii = 0
End If

End Sub
