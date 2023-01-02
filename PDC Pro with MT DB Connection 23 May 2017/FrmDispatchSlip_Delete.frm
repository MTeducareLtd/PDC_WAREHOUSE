VERSION 5.00
Begin VB.Form FrmDispatchSlip_Delete 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Dispatch Slip"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Accept"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
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
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   2415
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
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1875
   End
End
Attribute VB_Name = "FrmDispatchSlip_Delete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
'If checkvalid = False Then Exit Sub

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim str As String
    
str = "Select * from ASPDC_DispatchSlip where DispatchSlipCode ='" & Trim(UCase(Me.txtSlipNo.Text)) & "'"
rs1.Open str, cn1, adOpenDynamic, adLockReadOnly
If (rs1.BOF And rs1.EOF) Then
    MsgBox "Invalid Dispatch Slip Number.", vbCritical + vbOKOnly
    rs1.Close
    cn1.Close
    txtSlipNo.Text = ""
    txtSlipNo.SetFocus
    Exit Sub
End If
rs1.Close


'Check if any cheque from the selected slip is already sent to bank or not
rs1.Open "Select * from ASPDC_DispatchSlipDetails where DispatchSlipCode ='" & Trim(UCase(Me.txtSlipNo.Text)) & "' and CMSDoneflag =1", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    MsgBox "CMS is done for few cheques in this slip.  You can't Delete this Dispatch Slip.", vbCritical + vbOKOnly
    rs1.Close
    cn1.Close
    txtSlipNo.Text = ""
    txtSlipNo.SetFocus
    Exit Sub
End If
rs1.Close

'Delete entry
rs1.Open "Delete from ASPDC_DispatchSlipDetails where DispatchSlipCode ='" & Trim(UCase(Me.txtSlipNo.Text)) & "'", cn1, adOpenDynamic, adLockPessimistic
rs1.Open "Delete from ASPDC_DispatchSlip where DispatchSlipCode ='" & Trim(UCase(Me.txtSlipNo.Text)) & "'", cn1, adOpenDynamic, adLockPessimistic

cn1.Close

MsgBox "Dispatch Slip Entry removed successfully.", vbInformation + vbOKOnly
txtSlipNo.Text = ""
End Sub

Private Sub txtSlipNo_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cmdSave_Click
    KeyAscii = 0
End If
End Sub
