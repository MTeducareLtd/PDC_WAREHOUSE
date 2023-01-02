VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReturnBounceChequeEntry_Sub1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bounce Cheque Return Entry"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtChequeDate 
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
      Left            =   4440
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtCentreName 
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
      Left            =   3720
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtChqAmt 
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
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3000
      Width           =   1695
   End
   Begin VB.ComboBox cboReason 
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
      ItemData        =   "FrmReturnBounceChequeEntry_Sub1.frx":0000
      Left            =   3000
      List            =   "FrmReturnBounceChequeEntry_Sub1.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtCCChqNo 
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
      Top             =   720
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
      Left            =   3000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtChqBarcodeNo 
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
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Accept"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
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
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComCtl2.DTPicker dtEntry 
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
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
      Format          =   137822211
      CurrentDate     =   39310
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Returned To Location"
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
      TabIndex        =   11
      Top             =   1200
      Width           =   1875
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Date"
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
      TabIndex        =   10
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Barcode No"
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
      Top             =   240
      Width           =   1725
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
      TabIndex        =   8
      Top             =   2160
      Width           =   720
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Number"
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
      TabIndex        =   6
      Top             =   720
      Width           =   1365
   End
End
Attribute VB_Name = "FrmReturnBounceChequeEntry_Sub1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cboReason_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cmdSave_Click
    KeyAscii = 0
End If
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
Unload Me
End Sub


Private Sub cmdPrint_Click()
On Error GoTo ExitStep
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

Dim CentreCode As String
CentreCode = Me.txtCentreName.Text

'Validate barcode no and cheque no
If Len(Me.txtCentreName.Text) = 5 Then
    rs2.Open "select Source_Division_ShortDesc as divisionname, Source_Center_Name as centrename from C008_Centers g inner join C006_Division  d on left(g.Source_Center_Code,2) = d.source_division_code where g.Target_Center_Code ='" & Me.txtCentreName.Text & "'", cn1, adOpenDynamic, adLockReadOnly
Else
    rs2.Open "select Source_Division_ShortDesc as divisionname, Source_Center_Name as centrename from C008_Centers g inner join C006_Division  d on left(g.Source_Center_Code,2) = d.source_division_code where g.Source_Center_Code ='" & Me.txtCentreName.Text & "'", cn1, adOpenDynamic, adLockReadOnly
End If
If Not (rs2.BOF And rs2.EOF) Then
    Me.txtCentreName.Text = rs2!DivisionName & " - " & rs2!CentreName
End If
rs2.Close
cn1.Close

With FrmReturnBounceChequeEntry_Sub2
    .lblCentreName.Caption = Me.txtCentreName.Text
    .lblCentreName1.Caption = .lblCentreName.Caption
    
    .lblChequeAmt.Caption = Me.txtChqAmt.Text
    .lblChequeAmt1.Caption = .lblChequeAmt.Caption
    
    .lblChequeDate.Caption = Me.txtChequeDate.Text
    .lblChequeDate1.Caption = .lblChequeDate.Caption
    
    .lblChequeNo.Caption = Me.txtCCChqNo.Text
    .lblChequeNo1.Caption = .lblChequeNo.Caption
    
    .lblReason.Caption = Me.cboReason.Text
    .lblReason1.Caption = .lblReason.Caption
    
    .lblReturnDate.Caption = Format(Me.dtEntry.Value, "dd Mmm yyyy")
    .lblReturnDate1.Caption = .lblReturnDate.Caption
    
    .lblBarcode.Caption = Me.txtChqBarcodeNo.Text
    .lblcentrecode.Caption = CentreCode
    
    .lblBarcode1.Caption = "*" & Me.txtChqBarcodeNo.Text & "*"
    .lblBarcode2.Caption = "*" & Me.txtChqBarcodeNo.Text & "*"
    
    
    .PrintForm
End With
Unload FrmReturnBounceChequeEntry_Sub2
Exit Sub

ExitStep:
MsgBox Err.Description
End Sub

Private Sub cmdSave_Click()
On Error Resume Next

If checkvalid = False Then Exit Sub

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

'Validate barcode no and cheque no
str1 = "Select * from ASPDC_BounceChequeEntry asb inner join ASPDC_Dispatchslipdetails asd on asb.CCChqIdNo = asd.CCChqIdNo where asb.CCCHQIdNo ='" & Me.txtChqBarcodeNo.Text & "' and cast(asd.CCChequeNo as bigint) =" & Val(txtCCChqNo.Text) & " and (asd.CMSDoneFlag =1) and (asb.ReturnToCentreFlag =0 or asb.ReturnToCentreFlag is Null)"
rs1.Open str1, cn1, adOpenDynamic, adLockReadOnly
If (rs1.BOF And rs1.EOF) Then
    MsgBox "Invalid Cheque Barcode Number or Cheque Number", vbCritical + vbOKOnly
    txtChqBarcodeNo.SetFocus
    rs1.Close
    cn1.Close
    Exit Sub
Else
    Me.txtChqAmt.Text = Format(rs1!CCChequeAmt, "0.00")
    

        Me.txtCentreName.Text = Left(rs1!DispatchSlipCode, 5)

    
    Me.txtChequeDate.Text = Format(rs1!CCChequeDate, "dd Mmm yyyy")
End If
rs1.Close

'Check if any request is in pending status or not


Dim str As String
str1 = "Select * from ASPDC_BounceChequeEntry where CCCHQIdNo ='" & Me.txtChqBarcodeNo.Text & "'"
rs1.Open str1, cn1, adOpenDynamic, adLockPessimistic
If Not (rs1.BOF And rs1.EOF) Then
    rs1!chequelocation = Me.cboReason.Text
    rs1!ReturnToCentreFlag = 1
    rs1!ReturnToCentreDate = dtEntry.Value
    rs1.Update
End If
rs1.Close

cn1.Close

cmdSave.Enabled = False
cmdPrint.Enabled = True
cmdPrint.SetFocus
End Sub


Private Sub Form_Load()
On Error Resume Next
txtFlag.Text = "ADD"
dtEntry.Value = Date
TxtUserName.Text = ModInit.PDCUserName
cboReason.ListIndex = 0
End Sub


Private Function checkvalid() As Boolean
On Error Resume Next
If Trim(Me.txtChqBarcodeNo.Text) = "" Then
    MsgBox "Enter Cheque Barcode Number.", vbInformation + vbOKOnly
    txtChqBarcodeNo.SetFocus
    checkvalid = False
    Exit Function
End If

If Len(Trim(txtCCChqNo.Text)) = 0 Then
    MsgBox "Enter Cheque Number.", vbCritical + vbOKOnly
    txtCCChqNo.SetFocus
    checkvalid = False
    Exit Function
End If

If cboReason.ListIndex <= 0 Then
    MsgBox "Select Location where cheque is returned.", vbCritical + vbOKOnly
    cboReason.SetFocus
    checkvalid = False
    Exit Function
End If

checkvalid = True
End Function




Private Sub txtCCChqNo_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    Me.cboReason.SetFocus
    KeyAscii = 0
End If
End Sub

Private Sub txtChqBarcodeNo_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    Me.txtCCChqNo.SetFocus
    KeyAscii = 0
End If
End Sub
