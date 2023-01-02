VERSION 5.00
Begin VB.Form FrmDispatch_MICREntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheque - MICR Entry"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5220
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTranCode 
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
      Left            =   4200
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtMICRChqNo 
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
      Left            =   2520
      TabIndex        =   9
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtMICRScannerIn 
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
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtBarCode 
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
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox txtMICRNo 
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
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Accept"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Scanned Cheque Number"
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
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MICR Scanner Input"
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
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1740
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Bar Code"
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
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MICR Number - Tran Code"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   2250
   End
End
Attribute VB_Name = "FrmDispatch_MICREntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
'Validate
If Len(Trim(txtMICRNo.Text)) <> 9 Then
    MsgBox "Invalid MICR Number.", vbCritical + vbOKOnly
    Me.txtMICRNo.SetFocus
    Exit Sub
End If

'Save
On Error GoTo ErrExit
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

'Check for duplicate barcode
rs1.Open "Select * from ASPDC_DispatchSlipDetails where Location_Code ='" & ModInit.LocationCode & "' and CCChqIdNo ='" & Me.txtBarCode.Text & "'", cn1, adOpenDynamic, adLockPessimistic
If Not (rs1.BOF And rs1.EOF) Then
    'Check if this cheque number is matching with that read from scanner
    If Trim(Me.txtMICRChqNo.Text) <> "" And rs1!CCChequeNo <> Me.txtMICRChqNo.Text Then
        MsgBox "Invalid Cheque Number.  Cheque number entered in system for this cheque is " & rs1!CCChequeNo & ".", vbCritical + vbOKOnly
        rs1.Close
        cn1.Close
        Exit Sub
    End If
    
    'Save MICR Number
    rs1!MICRNumber = Trim(Me.txtMICRNo.Text)
    rs1!TranCode = Trim(Me.txtTranCode.Text)
    rs1.Update
Else
    MsgBox "Invalid Cheque Barcode.", vbCritical + vbOKOnly
   
End If
rs1.Close
cn1.Close

'Clear entry
Me.txtBarCode.Text = ""
Me.txtMICRNo.Text = ""
Me.txtMICRChqNo.Text = ""
Me.txtMICRScannerIn.Text = ""
Me.txtTranCode.Text = ""
Me.txtBarCode.SetFocus
Exit Sub

ErrExit:

MsgBox Err.Description
End Sub

Private Sub txtBarCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    txtMICRScannerIn.SetFocus
    KeyAscii = 0
End If
End Sub

Private Sub txtMICRNo_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    txtTranCode.SetFocus
    KeyAscii = 0
End If
End Sub

Private Sub txtMICRScannerIn_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    If Len(txtMICRScannerIn.Text) = 0 Then
        Me.txtMICRNo.SetFocus
    Else
        'Read MICR NUmber from the scanned input
        If InStr(1, txtMICRScannerIn.Text, "?") > 0 Then
            'Invalid String hence ignore it
            txtMICRScannerIn.SelStart = 0
            txtMICRScannerIn.SelLength = Len(txtMICRScannerIn.Text)
            txtMICRScannerIn.SetFocus
            KeyAscii = 0
            Exit Sub
        Else
            'Valid string hence read MICR number from that string
            Me.txtMICRNo.Text = Mid(txtMICRScannerIn.Text, 10, 9)
            Me.txtMICRChqNo.Text = Mid(txtMICRScannerIn.Text, 2, 6)
            Me.txtTranCode.Text = Right(txtMICRScannerIn.Text, 2)
            KeyAscii = 0
            cmdSave_Click
        End If
        
    End If
    KeyAscii = 0
End If
End Sub

Private Sub txtTranCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cmdSave_Click
    KeyAscii = 0
End If
End Sub
