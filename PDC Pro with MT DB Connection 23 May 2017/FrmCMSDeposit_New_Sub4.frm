VERSION 5.00
Begin VB.Form FrmCMSDeposit_New_Sub4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirm Remove"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassKey 
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
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Remove"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password"
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
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "FrmCMSDeposit_New_Sub4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdRemove_Click()
On Error Resume Next
If txtPassKey.Text <> "Pdc" Then
    MsgBox "Invalid Password", vbCritical + vbOKOnly
    Exit Sub
End If

'Remove Cheque from CMS
Dim res As Integer
res = MsgBox("You are about to remove selected cheque from CMS.  Do you want to proceed?", vbQuestion + vbYesNo)

If res = 7 Then Exit Sub

'Remove entry from CMS using barcode
Dim Barcode As String
Barcode = FrmCMSDeposit_New_Sub1.Grid.TextMatrix(FrmCMSDeposit_New_Sub1.Grid.RowSel, 5)

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

rs1.Open "Update ASPDC_DispatchSlipDetails set cms_verifyflag =0, ICICIBankDSNo ='', CMSDoneFlag =0, CMSSlipCode ='', CMSSlipNo ='' where CCChqIDNo ='" & Barcode & "'", cn1, adOpenDynamic, adLockPessimistic

cn1.Close

'Update Grid
With FrmCMSDeposit_New_Sub1
    .FillGrid .txtSlipNo.Text, Val(.txtCMSSlipTypeCode.Text)
End With

Unload Me
End Sub
