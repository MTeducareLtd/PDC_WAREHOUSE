VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReturnChequeEntry 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Return Cheque Entry"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14700
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   14700
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Export"
      Height          =   375
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print Entry"
      Height          =   375
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   375
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2730
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   4815
      _Version        =   393216
      Cols            =   9
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
      FormatString    =   $"FrmReturnChequeEntry.frx":0000
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
      CustomFormat    =   "dd MMM yyyy"
      Format          =   138543107
      CurrentDate     =   39310
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   315
      Left            =   6120
      TabIndex        =   1
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
      CustomFormat    =   "dd MMM yyyy"
      Format          =   137822211
      CurrentDate     =   39310
   End
   Begin VB.CommandButton cmdEntry 
      Caption         =   "&Return Entry"
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
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
      TabIndex        =   9
      Top             =   3720
      Width           =   1470
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Date From"
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
      TabIndex        =   6
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Date To"
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
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "FrmReturnChequeEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub cmdClose_Click()
On Error Resume Next
Unload Me
End Sub

Public Sub cmdEntry_Click()
On Error Resume Next
With FrmReturnChequeEntry_Sub1
    .Show vbModal
End With
End Sub

Private Sub cmdPrint_Click()
On Error Resume Next
If Grid.Rows = 1 Then Exit Sub

With FrmReturnChequeEntry_Sub1
    .txtChqBarcodeNo.Text = Grid.TextMatrix(Grid.RowSel, 0)
    .txtCCChqNo.Text = Grid.TextMatrix(Grid.RowSel, 1)
    .txtChqAmt.Text = Grid.TextMatrix(Grid.RowSel, 2)
    .cboReason.Text = Grid.TextMatrix(Grid.RowSel, 4)
    .dtEntry.Value = Grid.TextMatrix(Grid.RowSel, 3)
    .txtCentreName.Text = Grid.TextMatrix(Grid.RowSel, 5)
    .txtChequeDate.Text = Grid.TextMatrix(Grid.RowSel, 7)
    .txtRemarks.Text = Grid.TextMatrix(Grid.RowSel, 8)
    .cmdSave.Enabled = False
    .cmdPrint.Enabled = True
    .Show vbModal
End With
End Sub

Private Sub cmdShow_Click()
ReadValues
End Sub

Private Sub dtFrom_Change()
Grid.Rows = 1
End Sub

Private Sub dtTo_Change()
Grid.Rows = 1
End Sub

Private Sub Form_Load()
On Error Resume Next
dtFrom.Value = Date
dtTo.Value = Date
Grid.Rows = 1
Me.cmdEntry.Visible = True
Me.cmdPrint.Visible = True
End Sub

Private Sub ReadValues()
On Error GoTo ErrExit
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset


Grid.Rows = 1
rs1.Open "select Dispatchslipcode, CCChqIdNo, CCChequeNo, CCChequeAmt, ReturnDate, ReturnReason, ReturnEffectDownloadFlag, CentreChequeDate, ManualMapReason from ASPDC_DispatchSlipDetails ADS where ReturnFlag =1 and ReturnDate >='" & Format(dtFrom.Value, "dd Mmm yyyy") & "' and ReturnDate <='" & Format(dtTo.Value, "dd Mmm yyyy") & "' order by ReturnDate, CCCHQIdNo", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        Grid.Rows = Grid.Rows + 1
        Grid.TextMatrix(Grid.Rows - 1, 0) = rs1.Fields("CCChqIdNo").Value
        Grid.TextMatrix(Grid.Rows - 1, 1) = rs1!CCChequeNo
        Grid.TextMatrix(Grid.Rows - 1, 2) = Format(rs1!CCChequeAmt, "0.00")
        Grid.TextMatrix(Grid.Rows - 1, 3) = Format(rs1.Fields("ReturnDate").Value, "dd Mmm yyyy")
        Grid.TextMatrix(Grid.Rows - 1, 4) = rs1.Fields("ReturnReason").Value
        If IsNumeric(Left(rs1.Fields("DispatchSlipCode"), 1)) = True Then
            Grid.TextMatrix(Grid.Rows - 1, 5) = Left(rs1.Fields("Dispatchslipcode").Value, 5)
        Else
            Grid.TextMatrix(Grid.Rows - 1, 5) = Left(rs1.Fields("Dispatchslipcode").Value, 4)
        End If
        
        If rs1!ReturnEffectDownloadFlag = 1 Then
            Grid.TextMatrix(Grid.Rows - 1, 6) = "Downloaded"
        Else
            Grid.TextMatrix(Grid.Rows - 1, 6) = "Not Downloaded"
        End If
        Grid.TextMatrix(Grid.Rows - 1, 7) = Format(rs1.Fields("CentreChequeDate").Value, "dd Mmm yyyy")
        Grid.TextMatrix(Grid.Rows - 1, 8) = "" & rs1.Fields("ManualMapReason").Value
        rs1.MoveNext
     Loop
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
DefColWidth = (Grid.Width - 350) / (Grid.Cols - 1)
With Grid
    For Cnt = 0 To Grid.Cols - 2
        .ColWidth(Cnt) = DefColWidth
    Next
    .ColWidth(8) = 0
End With
End Sub

Private Sub Form_Resize()
On Error Resume Next
Grid.Width = Me.Width - Grid.Left - 360
Shape1.Width = Grid.Width
SetGridWidth
cmdClose.Left = Grid.Left + Grid.Width - cmdClose.Width
cmdExport.Left = cmdClose.Left - cmdExport.Width - 120
Grid.Height = Me.Height - Grid.Top - 1200
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

wrkProjectName = Me.Caption

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

ModGridToExcel.Head1 = "Entry Date From : " & Format(dtFrom.Value, "dd Mmm yyyy")
ModGridToExcel.Head2 = "Entry Date To : " & Format(dtTo.Value, "dd Mmm yyyy")

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
            ModGridToExcel.FieldData(RowCnt, ColCnt + 1) = .TextMatrix(RowCnt, ColCnt)
        Next
    Next
End With

'  Save the Grid as a File
If (SaveExcelWorksheet(wrkProjectName, wrkOutputFile, wrkErrorMessage) = False) Then
    MsgBox "Error in data transfer."
    Exit Sub
End If
End Sub
