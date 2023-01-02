VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMonthlyStockCheck 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Monthly Stock Check"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14700
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   14700
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtResult 
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
      Left            =   6120
      TabIndex        =   15
      Top             =   960
      Width           =   5295
   End
   Begin VB.CheckBox chkUpdateMICR 
      Caption         =   "Update MICR Number"
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
      Left            =   11640
      TabIndex        =   13
      Top             =   240
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdAccept 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Accept"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   960
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
      Left            =   2280
      TabIndex        =   9
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtCentreCode 
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
      Left            =   6120
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Export"
      Height          =   375
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Stock"
      Height          =   375
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3210
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5662
      _Version        =   393216
      Cols            =   8
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
      FormatString    =   "<Dispatch Slip No|<Cheque Barcode No|<CC Cheque No|<CC Cheque Amount|<CC Cheque Date|<Centre Cheque Date|<Micr No|<Found Status"
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
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
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
      CustomFormat    =   "MMM yyyy"
      Format          =   125960195
      CurrentDate     =   39310
   End
   Begin MSFlexGridLib.MSFlexGrid GridErr 
      Height          =   3210
      Left            =   10440
      TabIndex        =   14
      Top             =   1560
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5662
      _Version        =   393216
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
      FormatString    =   "<Barcode No|<Remarks"
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
      TabIndex        =   10
      Top             =   960
      Width           =   1725
   End
   Begin VB.Label lblCnt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Record Count : 0"
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
      Top             =   4920
      Width           =   1470
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Month"
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
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Centre Code"
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
      Height          =   315
      Left            =   4080
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "FrmMonthlyStockCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub cmdAccept_Click()
On Error Resume Next
Dim FoundFlag As Boolean
FoundFlag = False

'Check for the barcode in the grid
For Cnt = 1 To Grid.Rows - 1
    If Grid.TextMatrix(Cnt, 1) = Me.txtChqBarcodeNo.Text And Grid.TextMatrix(Cnt, 7) <> "Found" Then
        Grid.TextMatrix(Cnt, 7) = "Found"
        
        'Check if micr number is there or not
        If Grid.TextMatrix(Cnt, 6) = "" And Me.chkUpdateMICR.Value = vbChecked Then
            With FrmDispatch_MICREntry
                .txtBarCode.Text = Me.txtChqBarcodeNo.Text
                .Show vbModal
            End With
        End If
        FoundFlag = True
        txtResult.Text = "Found"
        Exit For
    End If
Next

If FoundFlag = False Then
    'Add barcode in Invalid Barcode Item List
    GridErr.Rows = GridErr.Rows + 1
    GridErr.TextMatrix(GridErr.Rows - 1, 0) = Me.txtChqBarcodeNo.Text
    txtResult.Text = CheckError(txtChqBarcodeNo.Text)
    GridErr.TextMatrix(GridErr.Rows - 1, 1) = txtResult.Text
    
End If

txtChqBarcodeNo.Text = ""
txtChqBarcodeNo.SetFocus
End Sub

Private Function CheckError(BarcodeNo As String) As String
On Error GoTo ErrExit
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

rs1.Open "Select * from ASPDC_DispatchSlipDetails where CCChqIdNo ='" & BarcodeNo & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    If Left(rs1!DispatchSlipCode, 5) <> Me.txtCentreCode.Text Then
        CheckError = "Cheque of other centre"      'Cheque of other centre
    ElseIf Val("" & rs1!ReturnFlag) = 1 Then
        CheckError = "Return entry on " & Format(rs1!ReturnDate, "dd Mmm yy")
    ElseIf Val("" & rs1!CMSDoneFlag) = 1 Then
        CheckError = "CMS done on " & Format(rs1!CMSSlipDate, "dd Mmm yy")
    ElseIf Val("" & rs1!HoldFlag) = 1 Then
        CheckError = "Cheque on Hold"
    Else
        CheckError = "Wrong Cheque Date"
    End If
Else
    CheckError = "Invalid Barcode"
End If
rs1.Close
cn1.Close
Exit Function

ErrExit:
MsgBox Err.Description
CheckError = ""
End Function

Private Sub cmdClear_Click()
On Error Resume Next
dtFrom.Enabled = True
Me.txtCentreCode.Enabled = True
Grid.Rows = 1
Me.cmdShow.Enabled = True
Me.txtChqBarcodeNo.Text = ""
Me.cmdAccept.Enabled = False
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
Unload Me
End Sub





Private Sub cmdShow_Click()
ReadValues
End Sub

Private Sub dtFrom_Change()
Grid.Rows = 1
GridErr.Rows = 1
End Sub

Private Sub Form_Load()
On Error Resume Next
dtFrom.Value = Date
Me.chkUpdateMICR.BackColor = Me.BackColor
Grid.Rows = 1
End Sub

Private Sub ReadValues()
On Error GoTo ErrExit
Grid.Rows = 1
GridErr.Rows = 1

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

Dim StartDate As String
Dim EndDate As String

StartDate = "1 " & Format(dtFrom.Value, "MMM yyyy")
EndDate = Format(DateAdd("m", 1, DateValue(StartDate)), "dd Mmm yyyy")

'<Dispatch Slip No|<Cheque Barcode No|<CC Cheque No|<CC Cheque Amount|<CC Cheque Date|<Centre Cheque Date|<Micr No|<Found Status

Grid.Rows = 1
rs1.Open "select * from ASPDC_DispatchSlipDetails ADS where left(DispatchSlipCode,5) = '" & Me.txtCentreCode.Text & "' and CCChequeDate >='" & Format(StartDate, "dd Mmm yyyy") & "' and CCChequeDate <'" & Format(EndDate, "dd Mmm yyyy") & "' and (cmsdoneflag =0 or cmsdoneflag is null) and ccchqidno is not null and (HoldFlag =0 or HoldFlag is Null) order by CCCHQIdNo", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        Grid.Rows = Grid.Rows + 1
        Grid.TextMatrix(Grid.Rows - 1, 0) = rs1.Fields("DispatchSlipCode").Value
        Grid.TextMatrix(Grid.Rows - 1, 1) = rs1.Fields("CCChqIdNo").Value
        Grid.TextMatrix(Grid.Rows - 1, 2) = rs1!CCChequeNo
        Grid.TextMatrix(Grid.Rows - 1, 3) = Format(rs1!CCChequeAmt, "0.00")
        
        Grid.TextMatrix(Grid.Rows - 1, 4) = Format(rs1.Fields("CCChequeDate").Value, "dd Mmm yyyy")
        Grid.TextMatrix(Grid.Rows - 1, 5) = Format(rs1.Fields("CentreChequeDate").Value, "dd Mmm yyyy")
        
        Grid.TextMatrix(Grid.Rows - 1, 6) = "" & rs1.Fields("MICRNumber").Value
        Grid.TextMatrix(Grid.Rows - 1, 7) = ""
        rs1.MoveNext
     Loop
     
     Me.dtFrom.Enabled = False
     Me.txtCentreCode.Enabled = False
     Me.cmdShow.Enabled = False
     Me.cmdAccept.Enabled = True
     Me.txtChqBarcodeNo.SetFocus
Else
    MsgBox "Records not found.", vbCritical + vbOKOnly
    
End If
rs1.Close
cn1.Close

Me.lblCnt.Caption = "Record Count : " & Grid.Rows - 1

Exit Sub

ErrExit:
MsgBox "Error : " & Err.Description
End Sub

Private Sub SetGridWidth()
On Error Resume Next
Dim DefColWidth As Long
DefColWidth = (Grid.Width - 350) / (Grid.Cols)
With Grid
    For Cnt = 0 To Grid.Cols - 1
        .ColWidth(Cnt) = DefColWidth
    Next
End With

With GridErr
    .ColWidth(0) = 1000
    .ColWidth(1) = .Width - .ColWidth(0) - 350
End With
End Sub

Private Sub Form_Resize()
On Error Resume Next
Grid.Width = Me.Width - Grid.Left - 360 - GridErr.Width - 240
GridErr.Left = Grid.Left + Grid.Width + 240
SetGridWidth
cmdClose.Left = GridErr.Left + GridErr.Width - cmdClose.Width
cmdExport.Left = cmdClose.Left - cmdExport.Width - 120
Grid.Height = Me.Height - Grid.Top - 1200
GridErr.Height = Grid.Height

cmdClose.Top = Grid.Top + Grid.Height + 120
cmdExport.Top = cmdClose.Top
lblCnt.Top = cmdClose.Top
End Sub




Private Sub cmdExport_Click()
On Error Resume Next
Dim wrkErrorMessage As String
Dim wrkOutputFile As String
Dim wrkProjectName As String
Dim ColCnt, RowCnt As Integer

wrkProjectName = Me.Caption & "_" & Me.txtCentreCode.Text

'Set Output File Name
'cmd.CancelError = True
'cmd.Flags = cdlOFNOverwritePrompt
'cmd.DefaultExt = "xls"
'cmd.Filter = "Excel Files|*.XLS|All files|*.*"
'cmd.ShowSave
'If Err.Number = cdlCancel Then
'    Exit Sub
'Else
'    wrkOutputFile = cmd.FileName
'End If

ModGridToExcel.Head1 = "Month : " & Format(dtFrom.Value, "Mmm yyyy")
ModGridToExcel.Head2 = "Centre Code : " & Me.txtCentreCode.Text

With Grid
    '  Load Field Names
    ModGridToExcel.NumberColumns = 60
    ReDim ModGridToExcel.FieldNames(1 To 60)
    For ColCnt = 0 To .Cols
        ModGridToExcel.FieldNames(ColCnt) = .TextArray(ColCnt - 1)
    Next
    
    '  Load Data Array
    ModGridToExcel.NumberRows = .Rows - 1
    ReDim ModGridToExcel.FieldData(1 To .Rows - 1, 1 To 60)
    For RowCnt = 1 To .Rows - 1
        For ColCnt = 0 To .Cols - 1
            ModGridToExcel.FieldData(RowCnt, ColCnt + 1) = "'" & .TextMatrix(RowCnt, ColCnt)
        Next
    Next
End With

'  Save the Grid as a File
If (SaveExcelWorksheet(wrkProjectName, wrkOutputFile, wrkErrorMessage) = False) Then
    MsgBox "Error in data transfer."
    Exit Sub
End If
End Sub

Private Sub txtChqBarcodeNo_KeyPress(KeyAscii As Integer)
On Error Resume Next
txtResult.Text = ""
If KeyAscii = 13 Then
    cmdAccept_Click
    KeyAscii = 0
End If
End Sub
