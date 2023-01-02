VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmDispatchSlipAuthorisation_Sub3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheque Mapping"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkReturnToCentre 
      Caption         =   "Return to Centre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   6120
      Width           =   2655
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "e"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox txtChqBarCode 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3120
      Width           =   2415
   End
   Begin VB.ListBox lstRemarks 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      ItemData        =   "FrmDispatchSlipAuthorisation_Sub3.frx":0000
      Left            =   2640
      List            =   "FrmDispatchSlipAuthorisation_Sub3.frx":0010
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   3600
      Width           =   4335
   End
   Begin VB.TextBox txtOpenItemNo 
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
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtChqIdNo 
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6480
      Visible         =   0   'False
      Width           =   855
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
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cboChqType 
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
      ItemData        =   "FrmDispatchSlipAuthorisation_Sub3.frx":0084
      Left            =   2640
      List            =   "FrmDispatchSlipAuthorisation_Sub3.frx":0091
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2640
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker dtChqDate 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd MMM yyyy"
      Format          =   127074307
      CurrentDate     =   41403
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Accept"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox txtCCRemarks 
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
      Height          =   1155
      Left            =   2640
      MaxLength       =   250
      TabIndex        =   7
      Top             =   4680
      Width           =   4335
   End
   Begin VB.TextBox txtCCChequeAmt 
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
      Left            =   2640
      TabIndex        =   1
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtCCChequeNo 
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
      Left            =   2640
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtChequeAmt 
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
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtChequeEntry 
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
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CC Barcode"
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
      TabIndex        =   26
      Top             =   3120
      Width           =   1020
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Wrong Entry Reason"
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
      TabIndex        =   25
      Top             =   3600
      Width           =   1770
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<-- Approved"
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
      Left            =   5280
      TabIndex        =   24
      Top             =   1680
      Width           =   1110
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<-- Approved"
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
      Left            =   5280
      TabIndex        =   23
      Top             =   1200
      Width           =   1110
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Centre Cheque Number"
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
      TabIndex        =   20
      Top             =   240
      Width           =   1980
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CC Cheque Type"
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
      TabIndex        =   18
      Top             =   2640
      Width           =   1440
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CC Cheque Date"
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
      TabIndex        =   17
      Top             =   2160
      Width           =   1425
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      TabIndex        =   16
      Top             =   4680
      Width           =   750
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CC Cheque Amount"
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
      TabIndex        =   15
      Top             =   1680
      Width           =   1650
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CC Cheque Number"
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
      TabIndex        =   14
      Top             =   1200
      Width           =   1665
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Centre Cheque Amount"
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
      TabIndex        =   13
      Top             =   720
      Width           =   1965
   End
End
Attribute VB_Name = "FrmDispatchSlipAuthorisation_Sub3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
FrmDispatchSlipAuthorisation_Sub2.MisMatchCloseFlag = False
Unload Me
End Sub

Private Sub cmdEdit_Click()
On Error Resume Next
Dim NewBarCode As String
NewBarCode = Me.txtChqBarCode.Text

NewBarCode = InputBox("Enter New Barcode.", , NewBarCode)

If Len(NewBarCode) <> 8 And Len(NewBarCode) <> 6 Then
    MsgBox "Invalid Barcode.", vbCritical + vbOKOnly
Else
    Me.txtChqBarCode.Text = NewBarCode
End If

End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrExit
If Len(txtChqBarCode.Text) <> 8 And Len(txtChqBarCode.Text) <> 6 Then
    MsgBox "Invalid Cheque Barcod Number.", vbCritical + vbOKOnly
    Exit Sub
End If

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

'Check for duplicate barcode
rs1.Open "Select * from ASPDC_DispatchSlipDetails where CCChqIdNo ='" & Me.txtChqBarCode.Text & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    MsgBox "Duplicate Cheque Barcode No.", vbCritical + vbOKOnly
    rs1.Close
    cn1.Close
    Exit Sub
End If
rs1.Close

'chkReturnToCentre.Value = vbUnchecked

'Save entry in ASPDC_DispatchSlipDetails table
rs1.Open "select * from ASPDC_DispatchSlipDetails where DispatchSlipCode ='" & Me.txtSlipNo.Text & "' and ChqIDNo ='" & Me.txtChqIdNo.Text & "'", cn1, adOpenDynamic, adLockPessimistic
If Not (rs1.BOF And rs1.EOF) Then
    rs1!CCChequeNo = txtCCChequeNo.Text
    rs1!CCChequeAmt = Me.txtCCChequeAmt.Text
    rs1!CCChequeDate = Me.dtChqDate.Value
    rs1!CCChqIdNo = Me.txtChqBarCode.Text
    rs1!ManualMapReason = Me.txtCCRemarks.Text
    If Me.lstRemarks.Selected(0) = True Then
        rs1!ManualR1Flag = 1
    Else
        rs1!ManualR1Flag = 0
    End If
    If Me.lstRemarks.Selected(1) = True Then
        rs1!ManualR2Flag = 1
    Else
        rs1!ManualR2Flag = 0
    End If
    If Me.lstRemarks.Selected(2) = True Then
        rs1!ManualR3Flag = 1
    Else
        rs1!ManualR3Flag = 0
    End If
    If Me.lstRemarks.Selected(3) = True Then
        rs1!ManualR4Flag = 1
    Else
        rs1!ManualR4Flag = 0
    End If
    rs1!ManualEntryDate = Date
    rs1!ManualUserCode = ModInit.PDCUserName
    rs1!EffectDownloadFlag = 0
    
    If Me.chkReturnToCentre.Value = vbChecked Then
        rs1!ReturnFlag = 1
        rs1!ReturnDate = Date
        rs1!ReturnReason = "Wrong Cheque"
        rs1!ManualMapReason = "Dispatch Slip: " & Me.txtSlipNo.Text & vbCrLf & "Cheque as per slip " & Me.txtChequeEntry.Text & ", Amt " & Me.txtChequeAmt.Text & vbCrLf & "Cheque details " & Me.txtCCChequeNo.Text & ", Amt " & Me.txtCCChequeAmt.Text
    End If
    
    rs1.Update
End If
rs1.Close
    
'Save entry in ASPDC_DispatchSlip_OpenItems table
If Me.txtChqBarCode.Enabled = False Then
    rs1.Open "select * from ASPDC_DispatchSlip_OpenItems where DispatchSlipCode ='" & Me.txtSlipNo.Text & "' and OpenItemEntryNo =" & Me.txtOpenItemNo.Text & "", cn1, adOpenDynamic, adLockPessimistic
    If Not (rs1.BOF And rs1.EOF) Then
        rs1!OpenEntryFlag = 0
        rs1!LinkedCHQIdNo = Me.txtChqIdNo.Text
        rs1.Update
    End If
    rs1.Close
End If

'Save entry in ASPDC_DispatchSlip table
rs1.Open "Select * from ASPDC_DispatchSlip where DispatchSlipCode ='" & Me.txtSlipNo.Text & "'", cn1, adOpenDynamic, adLockPessimistic
If Not (rs1.BOF And rs1.EOF) Then
    rs1!ManualMapChqCnt = Val(rs1!ManualMapChqCnt) + 1
    rs1!OpenChqCnt = Val(rs1!OpenChqCnt) - 1
    rs1.Update
End If
rs1.Close
cn1.Close


FrmDispatchSlipAuthorisation_Sub2.MisMatchCloseFlag = True

Unload Me
Exit Sub

ErrExit:
MsgBox Err.Description
End Sub

Private Sub txtCCChequeNo_KeyPress(KeyAscii As Integer)
On Error Resume Next
If IsNumeric(Asc(KeyAscii)) = True Then
    'Accept entry
ElseIf KeyAscii = 8 Then
    'Accept entry (backspace)
ElseIf KeyAscii = 13 Then
    txtCCChequeAmt.SetFocus
    KeyAscii = 0
Else
    KeyAscii = 0    'Non numeric value
End If
End Sub

