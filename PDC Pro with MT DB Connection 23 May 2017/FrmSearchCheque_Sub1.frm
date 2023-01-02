VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmSearchCheque_Sub1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Remove"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtChqBarCode 
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   1890
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3334
      _Version        =   393216
      Cols            =   4
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
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "<CMS Slip No|<CMS Slip Date|<Cheque Cnt|<Remove Flag"
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
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Barcode"
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
      TabIndex        =   2
      Top             =   240
      Width           =   1425
   End
End
Attribute VB_Name = "FrmSearchCheque_Sub1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub FillDuplicateCMS()
On Error GoTo ErrExit
Grid.Rows = 1
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

'Check for duplicate barcode
rs1.Open "Select ChqIDNo from ASPDC_DispatchSlipDetails where CCChqIdNo ='" & Me.txtChqBarCode.Text & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    Me.txtChqIdNo.Text = "" & rs1!ChqIDNo
    
    rs2.Open "select SlipNo, SlipDate, ChqCnt from Depositslipdetails DS inner join DepositSlip D on DS.SlipCode = D.SlipCode and DS.CentreCode = d.centrecode where ChequeIdNo ='" & Me.txtChqIdNo.Text & "'", cn1, adOpenDynamic, adLockReadOnly
    If Not (rs2.BOF And rs2.EOF) Then
        rs2.MoveFirst
        Do While Not rs2.EOF
            Grid.Rows = Grid.Rows + 1
            Grid.TextMatrix(Grid.Rows - 1, 0) = rs2!SlipNo
            Grid.TextMatrix(Grid.Rows - 1, 1) = Format(rs2!SlipDate, "dd Mmm yyyy")
            Grid.TextMatrix(Grid.Rows - 1, 2) = rs2!ChqCnt
            rs2.MoveNext
        Loop
    End If
    rs2.Close
End If
rs1.Close
cn1.Close
Exit Sub

ErrExit:
MsgBox Err.Description, vbCritical + vbOKOnly
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdRemove_Click()
On Error GoTo ErrExit
Dim CurRowNo As Integer
CurRowNo = 0
Dim RetainRowNo As Integer
For Cnt = 1 To Grid.Rows - 1
    If Grid.TextMatrix(Cnt, 3) = "Remove" Then
        CurRowNo = Cnt
        Exit For
    End If
Next

If CurRowNo = 0 Then
    MsgBox "Select Slip from which you want to remove Duplicate Cheque Entry.", vbCritical + vbOKOnly
    Exit Sub
End If

If CurRowNo = 1 Then
    RetainRowNo = 2
Else
    RetainRowNo = 1
End If

Dim res As Integer
res = MsgBox("You are about to remove duplicate entry from CMS Slip No. " & Grid.TextMatrix(CurRowNo, 0) & ".  Do you want to proceed now?", vbQuestion + vbYesNo)

Dim SlipCodeRemove As String
SlipCodeRemove = Grid.TextMatrix(CurRowNo, 0)

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

'Remove entry from Depositslipdetails
rs2.Open "delete from Depositslipdetails where slipcode ='" & SlipCodeRemove & "' and chequeidno ='" & txtChqIdNo.Text & "'", cn1, adOpenDynamic, adLockPessimistic

'Update ASPDC_DispatchSlipDetails
rs2.Open "update ASPDC_DispatchSlipDetails set CMSSlipCode ='" & Grid.TextMatrix(RetainRowNo, 0) & "', CMSSlipNo ='" & Grid.TextMatrix(RetainRowNo, 0) & "', CMSSlipDate ='" & Grid.TextMatrix(RetainRowNo, 1) & "'  where chqidno ='" & txtChqIdNo.Text & "'", cn1, adOpenDynamic, adLockPessimistic

'Change value in Depositslip for removed slip
rs1.Open "Select count(*) as FinalChqCnt, sum(Chequeamt) as FinalChqAmt from Depositslipdetails where SlipCode ='" & SlipCodeRemove & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs2.Open "Update ASPDC_CMS_CentreLog set FinalChequeAmt =" & rs1!FinalChqAmt & " , FinalChequeCnt =" & rs1!FinalChqCnt & " where CMSSlipNo ='" & SlipCodeRemove & "'", cn1, adOpenDynamic, adLockPessimistic
    rs2.Open "Update DepositSlip set chqcnt =" & rs1!FinalChqCnt & ", ChqAmt =" & rs1!FinalChqAmt & " where SlipCode =" & SlipCodeRemove, cn1, adOpenDynamic, adLockPessimistic
End If
rs1.Close

cn1.Close
MsgBox "Duplicate Entry removed.", vbInformation + vbOKOnly

FrmSearchCheque.cmdSave_Click
Unload Me
Exit Sub

ErrExit:
MsgBox Err.Description, vbInformation + vbOKOnly
End Sub

Private Sub Form_Load()
On Error Resume Next
With Grid
    .Rows = 1
    For Cnt = 0 To .Cols - 1
        .ColWidth(Cnt) = (.Width - 350) / .Cols
    Next
End With
End Sub

Private Sub Grid_DblClick()
On Error Resume Next
If Grid.Rows = 1 Then Exit Sub

Dim CurRowNo As Integer
CurRowNo = Grid.RowSel

For Cnt = 1 To Grid.Rows - 1
    Grid.TextMatrix(Cnt, 3) = ""
Next

Grid.TextMatrix(CurRowNo, 3) = "Remove"
End Sub
