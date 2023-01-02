VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmCMSDeposit_Modify 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify CMS Slip"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11175
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11175
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstBarcode 
      Height          =   1425
      Left            =   5760
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtChequeCnt 
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
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
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
      Left            =   11160
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "ADD"
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4770
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   8414
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      RowHeightMin    =   315
      BackColor       =   16761024
      ForeColor       =   0
      BackColorFixed  =   16744576
      ForeColorFixed  =   0
      BackColorSel    =   8388608
      ForeColorSel    =   16777215
      BackColorBkg    =   16761024
      GridColor       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"FrmCMSDeposit_Modify.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12000
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Count"
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
      TabIndex        =   9
      Top             =   240
      Width           =   1245
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CMS Slip Number"
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
      TabIndex        =   7
      Top             =   240
      Width           =   1485
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CMS Date"
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
      Left            =   2760
      TabIndex        =   6
      Top             =   240
      Width           =   870
   End
End
Attribute VB_Name = "FrmCMSDeposit_Modify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private Function CheckMissingCheque(ChqBarCode As String) As Boolean
'On Error GoTo ErrExit
'Dim cn1 As ADODB.Connection
'Set cn1 = New ADODB.Connection
'
'cn1.ConnectionString = ModInit.ConnectStringOnline
'cn1.Open
'
'Dim rs1 As ADODB.Recordset
'Set rs1 = New ADODB.Recordset
'
'Dim rs2 As ADODB.Recordset
'Set rs2 = New ADODB.Recordset
'
'Dim rs3 As ADODB.Recordset
'Set rs3 = New ADODB.Recordset
'
'Dim rs4 As ADODB.Recordset
'Set rs4 = New ADODB.Recordset
'
'Dim SerStr As String
'rs2.Open "Select * from ASPDC_DispatchSlipDetails where CCCHQIdNo ='" & ChqBarCode & "'", cn1, adOpenDynamic, adLockReadOnly
'If Not (rs2.BOF And rs2.EOF) Then
'    ChequeIdNo = rs2!ChqIDNo
'End If
'rs2.Close
'
'rs1.Open "Select * from DepositSlipDetails where SlipCode =" & Me.txtSlipNo.Text & " and ChequeIdNo ='" & ChequeIdNo & "'", cn1, adOpenDynamic, adLockReadOnly
'If Not (rs1.BOF And rs1.EOF) Then
'    rs2.Open "Select * from ASPDC_DispatchSlipDetails where CCCHQIdNo ='" & ChqBarCode & "'", cn1, adOpenDynamic, adLockPessimistic
'    If Not (rs2.BOF And rs2.EOF) Then
'        rs2!CMSSlipCode = txtSlipNo.Text
'        rs2!CMSSlipNo = txtSlipNo.Text
'        rs2!CMSSlipDate = Me.txtSlipDate.Text
'        rs2.Update
'    End If
'    rs2.Close
'    CheckMissingCheque = True
'Else
'    CheckMissingCheque = False
'End If
'rs1.Close
'cn1.Close
'Exit Function
'
'ErrExit:
'CheckMissingCheque = False
'End Function

Private Sub cmdCancel_Click()
On Error Resume Next
Unload Me
End Sub


Private Sub cmdSave_Click()
On Error GoTo ErrExit

'If checkvalid = False Then Exit Sub

Me.MousePointer = vbHourglass

'Save in Mirror table on local machine
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim ChqCnt As Integer
Dim ChqAmt As Double

For Cnt = 1 To Grid.Rows - 1
    If Grid.TextMatrix(Cnt, 0) = "" Then
        'Remove entry from depositslipdetails
        rs1.Open "delete from DepositSlipDetails where SlipCode =" & Me.txtSlipNo.Text & " and ChequeIdNo ='" & Grid.TextMatrix(Cnt, 13) & "'", cn1, adOpenDynamic, adLockPessimistic
                
        'Change cmsdone flag in aspdc
        rs1.Open "Update ASPDC_DispatchSlipDetails set cmsdoneflag =0, CMSSlipNo ='', CMSSlipCode =0 where ChqIDNo ='" & Grid.TextMatrix(Cnt, 13) & "' and CMSSlipNo ='" & txtSlipNo.Text & "'", cn1, adOpenDynamic, adLockPessimistic
    Else
        ChqCnt = ChqCnt + 1
        ChqAmt = ChqAmt + Val(Grid.TextMatrix(Cnt, 3))
    End If
Next

'change slip chq cnt and chq amt in depositslip tables
rs1.Open "Update DepositSlip set ChqCnt =" & ChqCnt & ", ChqAmt =" & ChqAmt & " where SlipNo ='" & Me.txtSlipNo.Text & "'", cn1, adOpenDynamic, adLockPessimistic

'change slip chq cnt and chq amt in aspdc_cms_log table
rs1.Open "Update ASPDC_CMS_CentreLog set FinalChequeCnt =" & ChqCnt & ", FinalChequeAmt =" & ChqAmt & " where CMSSlipNo ='" & Me.txtSlipNo.Text & "'", cn1, adOpenDynamic, adLockPessimistic


cn1.Close

MsgBox "CMS Slip modified successfully.", vbOKOnly + vbInformation
Unload Me
Exit Sub

ErrExit:
cn1.Close
MsgBox Err.Description
Me.MousePointer = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
txtFlag.Text = "ADD"
TxtUserName.Text = ModInit.PDCUserName
With Grid
    '<Deposit Flag|<Instrument No.|<Instrument Date|>Instrument Amount|<Bank Name|<Name of the Student|<Course|<Form No.|<Academic Year|<RcptCode|<SBEntryCode|<ChequeIdNo
    .ColWidth(0) = (.Width - 350) / (.Cols)
    .ColWidth(1) = .ColWidth(0)
    .ColWidth(2) = .ColWidth(0)
    .ColWidth(3) = .ColWidth(0)
    .ColWidth(4) = .ColWidth(0)
    .ColWidth(5) = .ColWidth(0)
    .Rows = 1
End With

End Sub

Public Sub FillGrid()
On Error GoTo ErrExit
If Trim(txtSlipNo.Text) = "" Then Exit Sub

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

Grid.Rows = 1

rs1.Open "select cmsdate, FinalChequeCnt from ASPDC_CMS_CentreLog AC where CMSSlipNo ='" & Me.txtSlipNo.Text & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    txtSlipDate.Text = Format(rs1.Fields("cmsdate").Value, "dd Mmm yyyy")
    txtChequeCnt.Text = rs1.Fields("FinalChequeCnt").Value
Else
    MsgBox "Invalid CMS Slip Number.", vbCritical + vbOKOnly
    txtSlipNo.SetFocus
    Exit Sub
End If
rs1.Close

Me.MousePointer = vbHourglass

Grid.Cols = 15
rs1.Open "Select * from DepositSlipDetails where SlipCode ='" & txtSlipNo.Text & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        Grid.Rows = Grid.Rows + 1
        
        '<Status|<Cheque No|<Cheque Date|<Cheque Amt|<Student Name|<Stream Name|<Form No|<Year Name|<Batch Code|<Roll No|<Rcpt Code|<Cheque Date|<Bank Name|<ChqId|<Bar Code
        
        Grid.TextMatrix(Grid.Rows - 1, 0) = "Deposited"
        Grid.TextMatrix(Grid.Rows - 1, 1) = rs1!ChequeNo
        Grid.TextMatrix(Grid.Rows - 1, 2) = Format(rs1.Fields("ChequeDate").Value, "ddMmyyyy")
        Grid.TextMatrix(Grid.Rows - 1, 3) = Format(rs1.Fields("ChequeAmt").Value, "0.00")
        
        
        Grid.TextMatrix(Grid.Rows - 1, 7) = rs1!YearName
        Grid.TextMatrix(Grid.Rows - 1, 8) = rs1!BatchCode
        Grid.TextMatrix(Grid.Rows - 1, 9) = rs1!RollNo
        Grid.TextMatrix(Grid.Rows - 1, 10) = rs1!RcptCode

        Grid.TextMatrix(Grid.Rows - 1, 11) = Format(rs1!ChequeDate, "dd MMM yyyy")


        rs3.Open "Select Title, FirstName, MidName, LastName, Sex, SBEntryCode from Student inner join StudentBatch on Student.Yearname = StudentBatch.YearName and Student.InstituteCode = StudentBatch.InstituteCode and Student.LicCode = StudentBatch.LicCode and Student.StudentCode = StudentBatch.StudentCode where StudentBatch.RecordDelFlag = 0 and StudentBatch.BatchCode ='" & rs1.Fields("BatchCode").Value & "' and StudentBatch.RollNo =" & rs1.Fields("RollNo").Value, cn1, adOpenDynamic, adLockReadOnly
        If Not (rs3.BOF And rs3.EOF) Then
            Grid.TextMatrix(Grid.Rows - 1, 4) = rs3.Fields("FirstName").Value & " " & Left(rs3.Fields("MidName").Value, 1) & " " & rs3.Fields("LastName").Value
            Grid.TextMatrix(Grid.Rows - 1, 6) = rs3!Sex
        End If
        rs3.Close
        
        rs3.Open "Select BankName from StudentPayment where SBEntryCode ='" & rs1!SBEntryCode & "' and RcptCode ='" & rs1!RcptCode & "'", cn1, adOpenDynamic, adLockReadOnly
        If Not (rs3.BOF And rs3.EOF) Then
            Grid.TextMatrix(Grid.Rows - 1, 12) = "" & rs3!BankName
        End If
        rs3.Close

        rs3.Open "SELECT Streams.StreamName FROM StudentBatch INNER JOIN (Batches INNER JOIN Streams ON Batches.StreamCode = Streams.StreamCode) ON StudentBatch.BatchCode = Batches.BatchCode WHERE  StudentBatch.RecordDelFlag = 0 and StudentBatch.BatchCode='" & rs1.Fields("BatchCode").Value & "' AND StudentBatch.RollNo=" & rs1.Fields("RollNo").Value, cn1, adOpenDynamic, adLockReadOnly
        If Not (rs3.BOF And rs3.EOF) Then
            Grid.TextMatrix(Grid.Rows - 1, 5) = rs3!StreamName
        End If
        rs3.Close
        
        Grid.TextMatrix(Grid.Rows - 1, 13) = rs1!ChequeIdNo
        
        rs3.Open "Select CCCHQIdNo from ASPDC_DispatchSlipDetails where CHQIdNo ='" & rs1!ChequeIdNo & "'", cn1, adOpenDynamic, adLockReadOnly
        If Not (rs3.BOF And rs3.EOF) Then
            Grid.TextMatrix(Grid.Rows - 1, 14) = rs3!CCCHQIdNo
        End If
        rs3.Close
        
        Err.Clear
        rs1.MoveNext
        If Err.Number > 0 Then GoTo ErrExit
    Loop
End If
rs1.Close

cn1.Close

Me.MousePointer = 0
Exit Sub

ErrExit:
Me.MousePointer = 0
MsgBox Err.Description
End Sub

Private Function checkvalid() As Boolean
On Error Resume Next
If Trim(Me.txtSlipNo.Text) = "" Then
    MsgBox "Enter CMS Slip Number.", vbInformation + vbOKOnly
    txtSlipNo.SetFocus
    checkvalid = False
    Exit Function
End If

If Val(Me.txtChequeCnt.Text) <> Grid.Rows - 1 Then
    MsgBox "You have not entered details of all cheques in the slip.", vbCritical + vbOKOnly
    Grid.SetFocus
    checkvalid = False
    Exit Function
End If

checkvalid = True
End Function


Private Sub Grid_DblClick()
On Error Resume Next
If Grid.Rows = 1 Then Exit Sub

If Grid.TextMatrix(Grid.RowSel, 0) = "Deposited" Then
    Grid.TextMatrix(Grid.RowSel, 0) = ""
Else
    Grid.TextMatrix(Grid.RowSel, 0) = "Deposited"
End If

Dim ChqCnt As Integer
For Cnt = 1 To Grid.Rows - 1
    If Grid.TextMatrix(Cnt, 0) = "Deposited" Then
        ChqCnt = ChqCnt + 1
    End If
Next

Me.txtChequeCnt.Text = ChqCnt
End Sub

