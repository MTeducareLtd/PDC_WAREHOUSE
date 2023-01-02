VERSION 5.00
Begin VB.Form FrmCMSDeposit_New_Sub3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add 2 CMS"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5730
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Accept"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
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
      TabIndex        =   5
      Top             =   720
      Width           =   1365
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
      TabIndex        =   4
      Top             =   240
      Width           =   1725
   End
End
Attribute VB_Name = "FrmCMSDeposit_New_Sub3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
'validation
On Error GoTo ErrExit
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

Dim rs3 As ADODB.Recordset
Set rs3 = New ADODB.Recordset

Dim rs4 As ADODB.Recordset
Set rs4 = New ADODB.Recordset

Dim ChequeAmt As Single
Dim ChqIdNo As String
rs1.Open "Select * from aspdc_dispatchslipdetails where CCChqIdNo ='" & Me.txtChqBarcodeNo.Text & "' and CCChequeNo ='" & Me.txtCCChqNo.Text & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    'check if cms for the cheque is already generated or not
    If Val("" & rs1!CMSDoneFlag) = 1 Then
        MsgBox "CMS for this cheque is already done.", vbCritical + vbOKOnly
        rs1.Close
        cn1.Close
        Exit Sub
    End If
    
    'check if cheque is on hold
    If Val("" & rs1!HoldFlag) = 1 Then
        MsgBox "This cheque is on Hold hence can't be added in CMS", vbCritical + vbOKOnly
        rs1.Close
        cn1.Close
        Exit Sub
    End If
    
    'check if cheque is returned to centre
    If Val("" & rs1!ReturnFlag) = 1 Then
        MsgBox "This cheque is Returned to Centre hence can't be added in CMS", vbCritical + vbOKOnly
        rs1.Close
        cn1.Close
        Exit Sub
    End If
    
    'check if cheque date is current date or past date
    If rs1!CCChequeDate > FrmCMSDeposit_New.dtSlip.Value Then
        MsgBox "This is a Future date cheque hence can't be added in CMS", vbCritical + vbOKOnly
        rs1.Close
        cn1.Close
        Exit Sub
    End If
    
    ChqIdNo = "" & rs1!ChqIdNo
Else
    MsgBox "Invalid Barcode or Cheque Number", vbCritical + vbOKOnly
    rs1.Close
    cn1.Close
    Exit Sub
End If
rs1.Close

'check if admission is cancelled for the student
Dim SBEntryCode, RcptCode As String
Dim ChequeDate As Date

'rs1.Open "Select T.SBEntryCode, T.Status, T.RecordDelFlag, T.PendingFlag, SP.RecordDelFlag as SPRecordDelFlag, SP.ChkDate, SP.RcptCode from StudentBatch T inner join StudentPayment SP on T.Sbentrycode = sp.sbentrycode where sp.ChequeIdNo ='" & ChqIdNo & "'", cn1, adOpenDynamic, adLockReadOnly
'If Not (rs1.BOF And rs1.EOF) Then
'    SBEntryCode = rs1!SBEntryCode
'    ChequeDate = rs1!ChkDate
'    RcptCode = rs1!RcptCode
'
'    If rs1!Status = 0 And rs1!PendingFlag = 0 Then
'        MsgBox "Admission is cancelled for the Student hence cheque can't be added in CMS", vbCritical + vbOKOnly
'        rs1.Close
'        cn1.Close
'        Exit Sub
'    End If
'
'    If rs1!RecordDelFlag = 1 Or rs1!SPRecordDelFlag = 1 Then
'        MsgBox "Cheque is deleted for the Student hence cheque can't be added in CMS", vbCritical + vbOKOnly
'        rs1.Close
'        cn1.Close
'        Exit Sub
'    End If
'Else
'    MsgBox "Student details not found for the Cheque Entry.", vbCritical + vbOKOnly
'    rs1.Close
'    cn1.Close
'    Exit Sub
'End If
'rs1.Close

'update entry in ASPDC table
Dim SlipNo, SAPCentreCode As String
SlipNo = Format(FrmCMSDeposit_New.dtSlip.Value, "ddMMyyyy")
rs1.Open "Select * from aspdc_dispatchslipdetails where CCChqIdNo ='" & Me.txtChqBarcodeNo.Text & "' and CCChequeNo ='" & Me.txtCCChqNo.Text & "'", cn1, adOpenDynamic, adLockPessimistic
If Not (rs1.BOF And rs1.EOF) Then
    rs1!CMSDoneFlag = 1
    rs1!CMSSlipCode = SlipNo
    rs1!CMSSlipNo = SlipNo
    rs1!CMSSlipDate = FrmCMSDeposit_New.dtSlip.Value
    
    ChequeAmt = rs1!CCChequeAmt
    
    'add entry in T005 table
'    rs3.Open "Select * from T005_CMS_Data where CMS_SLIPNO ='" & SlipNo & "' and Pay_InsNum ='" & txtCCChqNo.Text & "' and Cur_SB_code ='" & SBEntryCode & "'", cn1, adOpenDynamic, adLockPessimistic
'    If (rs3.BOF And rs3.EOF) Then
'        rs3.AddNew
'
'        SAPCentreCode = ""
'        rs4.Open "select * from C008_Centers where Target_Center_Code ='" & Left(SBEntryCode, 5) & "'", cn1, adOpenDynamic, adLockPessimistic
'        If Not (rs4.BOF And rs4.EOF) Then
'            rs3!Center_Code = rs4!Source_Center_Code
'            SAPCentreCode = rs4!Source_Center_Code
'        End If
'        rs4.Close
'
'        rs3!Cal_Year = Format(FrmCMSDeposit_New.dtSlip.Value, "YYYY")
'        rs3!Event_Type = "E14"
'        rs3!Event_Date = Format(FrmCMSDeposit_New.dtSlip.Value, "dd Mmm YYYY")
'        rs3!RECORD_ID = "S"
'        rs3!cms_slipno = SlipNo
'        rs3!CMS_SLIPDAT = FrmCMSDeposit_New.dtSlip.Value
'        rs3!Cur_SB_Code = SBEntryCode
'        rs3!PAY_MODE = "01"
'        rs3!pay_insnum = txtCCChqNo.Text
'        rs3!Pay_date = ChequeDate
'        rs3!Amt_instr = ChequeAmt
'        rs3!NARRATION1 = ""
'        rs3!Narration2 = ""
'        rs3!TRANS_FRMAPP = "01"
'        rs3!TRANS_FRMAREA = "01"
'        rs3!TRANS_FRMDAT = FrmCMSDeposit_New.dtSlip.Value
''        rs3!TRANS_FRMTIM = ""
''        rs3!TRANS_TOAPP = ""
''        rs3!TRANS_TOAREA = ""
''        rs3!TRANS_TODAT = ""
''        rs3!TRANS_TOTIM = ""
'        rs3!rec_Statid = "N"
'        rs3!RUN_NUMBER = ""
'
'        Dim BatchCode As String
'        BatchCode = ""
'        rs4.Open "Select AdmnDate, BatchCode from StudentBatch where SBEntryCode ='" & SBEntryCode & "'", cn1, adOpenDynamic, adLockReadOnly
'        If Not (rs4.BOF And rs4.EOF) Then
'            rs3!ADM_DATE = rs4!AdmnDate
'            BatchCode = rs4!BatchCode
'        End If
'        rs4.Close
'
'
'        rs4.Open "Select top 1 * from t001_admission_header where cur_sb_code  ='" & SBEntryCode & "' and Event_Type in ('E01' ,'P01') order by Event_Date desc", cn1, adOpenDynamic, adLockReadOnly
'        If Not (rs4.BOF And rs4.EOF) Then
'           rs3!Narration2 = rs4!Stream_Desc
'        End If
'        rs4.Close
'
'        NewNumber = 0
'        NewNumberStr = ""
'
'        rs4.Open "select isnull(max(right(rec_num,6)),0) as Last_Number_Used from T005_CMS_Data where center_code='" & SAPCentreCode & "' and cal_Year ='" & Format(FrmCMSDeposit_New.dtSlip.Value, "YYYY") & "' and right(rec_num,6) > 80000 and REC_NUM <> 'V'", cn1, adOpenDynamic, adLockReadOnly
'        If Not (rs4.BOF And rs4.EOF) Then
'            NewNumber = Val("" & rs4!Last_Number_Used) + 1
'
'            If NewNumber = 1 Then NewNumber = 80001
'
'
'            NewNumberStr = "T005" & Format(NewNumber, "000000")
'        Else
'            NewNumberStr = "T005" & Format(80001, "000000")
'        End If
'        rs4.Close
'
'        rs3!Rec_Num = NewNumberStr
'        rs3!TRANS_FRMDAT = FrmCMSDeposit_New.dtSlip.Value
'        rs3!TRANS_FRMTIM = Time
'        rs3.Update
'    End If
'    rs3.Close
    
    

    
    
    rs1!CMSSBEntryCode = ""
    rs1!cmscenter_Code = ""
    rs1!CMSCal_Year = Format(FrmCMSDeposit_New.dtSlip.Value, "YYYY")
    rs1!CMSRec_Num = ""
    rs1!CMSRece_Num = ""
    
    rs1.Update
End If
rs1.Close
cn1.Close

txtChqBarcodeNo.Text = ""
txtCCChqNo.Text = ""

txtChqBarcodeNo.SetFocus
Exit Sub

ErrExit:
MsgBox Err.Description
End Sub
