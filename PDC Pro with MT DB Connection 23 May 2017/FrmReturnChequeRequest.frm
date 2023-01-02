VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmReturnChequeRequest 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Return Cheque Request"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   14595
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Export"
      Height          =   375
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2730
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   4815
      _Version        =   393216
      Cols            =   10
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
      FormatString    =   $"FrmReturnChequeRequest.frx":0000
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
      TabIndex        =   3
      Top             =   3720
      Width           =   1470
   End
End
Attribute VB_Name = "FrmReturnChequeRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub cmdClose_Click()
On Error Resume Next
Unload Me
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

Private Sub cmdShow_Click()
ReadValues
End Sub


Private Sub Form_Load()
On Error Resume Next
Grid.Rows = 1
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

Dim SerStr As String
SerStr = "Select ASR.ReturnRequestCode, ASR.InstituteCode, ASR.LicCode, ASR.CenterChequeNo , ASR.CentreChequeAmt, ASP.CCChequeNo, ASP.CCChqIdNo, ASP.CCChequeDate, ASR.RequestApproveBy,  ASR.RequestApproveDate, ASR.CentreChequeDate ,ASR.RequestDate, Source_Center_Name as CenterName, Source_Division_ShortDesc as DivisionName, StudentName,  asr.SBEntryCode, ASP.CMSDoneFlag  " & _
         "from ASPDC_ReturnChequeRequest ASR inner join ASPDC_DispatchSlipDetails ASP on ASR.ChequeIdNo = ASP.ChqIdNo inner join C008_Centers C008 on C008.Source_Center_Code = ASR.InstituteCode + ASR.LicCode inner join C006_Division C006 on C006.Source_Division_Code = C008.Source_Division_Code " & _
         "where requestapproveFlag =1 and (ASP.ReturnFlag =0 or ASP.ReturnFlag is Null) order by ASR.InstituteCode, ASR.LicCode, CentreChequeDate, CenterChequeNo"

Grid.Rows = 1
rs1.Open SerStr, cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        If Val("" & rs1!CMSDoneFlag) = 0 Then
            '<Division|<Centre Name|<Centre Code|<Cheque Barcode No|<CC Cheque No|<CC Cheque Amount|<CC Cheque Date|<Return Request Date|<Request Approved By |<Request Approved On
            Grid.Rows = Grid.Rows + 1
            Grid.TextMatrix(Grid.Rows - 1, 0) = "" & rs1.Fields("DivisionName").Value
            Grid.TextMatrix(Grid.Rows - 1, 1) = "" & rs1!CenterName
            Grid.TextMatrix(Grid.Rows - 1, 2) = rs1!InstituteCode & rs1!LicCode
            Grid.TextMatrix(Grid.Rows - 1, 3) = "" & rs1!CCChqIdNo
            Grid.TextMatrix(Grid.Rows - 1, 4) = "" & rs1!CCChequeNo
            Grid.TextMatrix(Grid.Rows - 1, 5) = "" & Format(rs1.Fields("CentreChequeAmt").Value, "0.00")
            Grid.TextMatrix(Grid.Rows - 1, 6) = "" & Format(rs1.Fields("CCChequeDate").Value, "dd Mmm yyyy")
            Grid.TextMatrix(Grid.Rows - 1, 7) = "" & Format(rs1.Fields("RequestDate").Value, "dd Mmm yyyy")
            Grid.TextMatrix(Grid.Rows - 1, 8) = "" & rs1!RequestApproveBy
            Grid.TextMatrix(Grid.Rows - 1, 9) = "" & Format(rs1.Fields("RequestApproveDate").Value, "dd Mmm yyyy")
        End If
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
DefColWidth = (Grid.Width - 350) / (Grid.Cols)
With Grid
    For Cnt = 0 To Grid.Cols - 1
        .ColWidth(Cnt) = DefColWidth
    Next
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


