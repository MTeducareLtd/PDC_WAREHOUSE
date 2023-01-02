VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBounceChequeEntry_Sub1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bounce Cheque Entry"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboLocation 
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
      ItemData        =   "FrmBounceChequeEntry_Sub1.frx":0000
      Left            =   3000
      List            =   "FrmBounceChequeEntry_Sub1.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2640
      Width           =   2415
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
      ItemData        =   "FrmBounceChequeEntry_Sub1.frx":0028
      Left            =   3000
      List            =   "FrmBounceChequeEntry_Sub1.frx":0035
      Style           =   2  'Dropdown List
      TabIndex        =   2
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
      TabIndex        =   5
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
      TabIndex        =   7
      Top             =   3240
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
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
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "ADD"
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComCtl2.DTPicker dtEntry 
      Height          =   315
      Left            =   3000
      TabIndex        =   4
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
      Format          =   103350275
      CurrentDate     =   39310
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Location"
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
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bounce Reason"
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
      Top             =   1200
      Width           =   1365
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   8
      Top             =   720
      Width           =   1365
   End
End
Attribute VB_Name = "FrmBounceChequeEntry_Sub1"
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


Private Sub cmdSave_Click()
On Error Resume Next

If checkvalid = False Then Exit Sub

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

'Validate barcode no and cheque no
str1 = "Select * from ASPDC_DispatchSlipDetails where CCCHQIdNo ='" & Me.txtChqBarcodeNo.Text & "' and cast(CCChequeNo as bigint) =" & Val(txtCCChqNo.Text) & " and CMSDoneFlag =1"
rs1.Open str1, cn1, adOpenDynamic, adLockReadOnly
If (rs1.BOF And rs1.EOF) Then
    MsgBox "Invalid Cheque Barcode Number or Cheque Number", vbCritical + vbOKOnly
    txtChqBarcodeNo.SetFocus
    rs1.Close
    cn1.Close
    Exit Sub
End If
rs1.Close

'Check if bounce cheque entry already exists for this barcode no
rs1.Open "Select * from ASPDC_BounceChequeEntry where CCCHQIdNo ='" & txtChqBarcodeNo.Text & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    MsgBox "Bounce Cheque entry already exists for this Barcode Number.", vbCritical + vbOKOnly
    rs1.Close
    cn1.Close
    Exit Sub
End If
rs1.Close

Dim str As String
'Generate new BounceEntryCode
Dim CurDateStr As String
Dim LastBECode As String

CurDateStr = Format(Date, "ddMMyyyy")
rs1.Open "Select max(cast(right(BounceEntryCode,4) as int)) as LastBECode from  ASPDC_BounceChequeEntry where left(BounceEntryCode,8) ='" & CurDateStr & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    LastBECode = Val("" & rs1!LastBECode)
End If
rs1.Close
    
LastBECode = CurDateStr & Format(Val(LastBECode) + 1, "0000")
    
'Add entry in ASPDC_BounceChequeEntry table
rs1.Open "Select * from ASPDC_BounceChequeEntry where BounceEntryCode ='" & LastBECode & "'", cn1, adOpenDynamic, adLockPessimistic
rs1.AddNew
rs1!BounceEntryCode = LastBECode
rs1!CCCHQIdNo = Me.txtChqBarcodeNo.Text
rs1!BounceEntryDate = Date
rs1!BounceUserCode = Me.TxtUserName.Text
rs1!EffectDownloadFlag = 0
If Me.cboReason.ListIndex = 1 Then
    rs1!BouncePenaltyFlag = 1
Else
    rs1!BouncePenaltyFlag = 0
End If
rs1!ChequeLocation = Me.cboLocation.Text
rs1!Location_Code = ModInit.LocationCode
rs1!hideentry = 0
rs1.Update
rs1.Close

cn1.Close
Unload Me
FrmBounceChequeEntry.cmdEntry_Click
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
rs1.Fields("Location_Code").Value = ModInit.LocationCode
rs1.Update
rs1.Close

MsgBox "Dispatch Slip Cancellation Entry successfully done.", vbOKOnly + vbInformation
Unload Me
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
    MsgBox "Select Bounce Reason.", vbCritical + vbOKOnly
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
