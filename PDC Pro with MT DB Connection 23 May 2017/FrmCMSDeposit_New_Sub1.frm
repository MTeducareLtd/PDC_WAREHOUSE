VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmCMSDeposit_New_Sub1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Display CMS Slip"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   11115
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSMS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "S&MS"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6120
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox txtCTS 
      Height          =   2415
      Left            =   1080
      TabIndex        =   20
      Top             =   2520
      Visible         =   0   'False
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4260
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"FrmCMSDeposit_New_Sub1.frx":0000
   End
   Begin VB.TextBox txtICICIDSNo 
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
      Left            =   8880
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   480
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid BankExportGrid 
      Height          =   3255
      Left            =   480
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   1
      Cols            =   20
      FixedRows       =   0
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
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
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
   Begin VB.CommandButton cmdVerify 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Verify"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Export"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   840
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtCMSSlipTypeCode 
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
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox cboCompany 
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
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtCMSSlipType 
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
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   480
      Width           =   1935
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
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1335
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
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   1575
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
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   1575
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
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "ADD"
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4650
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8202
      _Version        =   393216
      Cols            =   16
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
      FormatString    =   $"FrmCMSDeposit_New_Sub1.frx":0082
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
   Begin VB.Label lblSMSCount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2160
      TabIndex        =   22
      Top             =   6120
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ICICI DS No"
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
      Left            =   8880
      TabIndex        =   19
      Top             =   240
      Width           =   1050
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Company Code"
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
      Left            =   6360
      TabIndex        =   14
      Top             =   240
      Width           =   1275
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CMS Slip Type"
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
      Left            =   4320
      TabIndex        =   13
      Top             =   240
      Width           =   1260
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
      Left            =   240
      TabIndex        =   11
      Top             =   6000
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
      TabIndex        =   9
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
      Left            =   2280
      TabIndex        =   8
      Top             =   240
      Width           =   870
   End
   Begin VB.Menu mnuExportForBank 
      Caption         =   "Options"
      Begin VB.Menu mnuProfundUploadFile 
         Caption         =   "Generate Bank Upload File"
      End
      Begin VB.Menu mnuSendSMS 
         Caption         =   "Send SMS"
      End
      Begin VB.Menu mnuCTSUploadFile 
         Caption         =   "CTS_Upload File"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFeedFile 
         Caption         =   "Feed File"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmCMSDeposit_New_Sub1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cboCompany_Click()
On Error Resume Next
FillGrid Me.txtSlipNo.Text, Val(txtCMSSlipTypeCode.Text)
End Sub

Private Sub cmdCancel_Click()
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

ModGridToExcel.Head1 = "Company : " & Me.cboCompany.Text
ModGridToExcel.Head2 = "CMS Slip Type : " & Me.txtCMSSlipType.Text
ModGridToExcel.Head3 = "CMS Slip No : " & Me.txtSlipNo.Text

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

Private Sub cmdPrint_Click()
On Error Resume Next
If txtCMSSlipTypeCode.Text = "4" Then
    MsgBox "Print option is not allowed for this type.", vbCritical + vbOKOnly
    Exit Sub
End If

On Error GoTo ErrHandler
cmd.CancelError = True

On Error GoTo ErrHandler
Dim NumCopies As Integer

'Set flags
cmd.Flags = &H100000 Or &H4

' Display the Print dialog box.
cmd.ShowPrinter
NumCopies = cmd.Copies
Printer.Orientation = cdlPortrait

PrintRes 1, 1

ErrHandler:
End Sub


Private Sub PrintRes(RowNo As Integer, NumCopies As Integer)
On Error GoTo ErrHandler
Dim TROWS, DPAGES, IPAGES, TPAGES, ENDNUM As Long
Dim intCopies  As Integer

Dim X As Integer
Dim Y As Integer
Dim nxtrow As Integer
Dim LeftMargin As Integer
Dim RightMargin As Long
Dim TopMargin As Long
Dim BottomMargin As Long
Dim YPos, YStart As Long
Dim NoOfCopies As Long
Dim SlipAmtTotal As Double

Dim RecordCnt As Integer
RecordCnt = Grid.Rows - 1

If RecordCnt = 0 Then Exit Sub

'For A4 size paper
LeftMargin = 2 * 567
RightMargin = 11000
TopMargin = 400
BottomMargin = 15000
SlipAmtTotal = 0

PrintLines = 40

' Column headings and blank line
TROWS = Grid.Rows - 1

'30 lines per sheet
DPAGES = TROWS / PrintLines

IPAGES = Int(DPAGES)

'Determine Number of Pages for Print
If DPAGES - IPAGES = 0 Then
    TPAGES = DPAGES
Else
    TPAGES = IPAGES + 1
End If

'If OptLocal.Value = True Then
    NoOfCopies = 1
'Else
'    NoOfCopies = 3
'End If

'Setup Global for printer
Printer.Font = "arial"
Printer.DrawWidth = 2
Printer.FillStyle = vbFSSolid
Printer.Orientation = cdlPortrait

Dim Col1Left, Col2Left, Col3Left, Col4Left, Col5Left, Col6Left, Col7Left, Col8Left, Col9Left, Col10Left As Long

'Start copies loop
For intCopies = 1 To NoOfCopies
    'Begin loop for pages

    For X = 1 To TPAGES
        YPos = TopMargin

        Printer.FontBold = True
        Printer.FontSize = 10
        Printer.FontName = "Arial"

        Printer.Line (LeftMargin, YPos)-(RightMargin, YPos)
        YStart = YPos
        YPos = YPos + 100
'        If OptLocal.Value = True Then
            PrintMatter = "ICICI BANK DEPOSIT SLIP"
'        Else
'            PrintMatter = "CITICHECK DEPOSIT SLIP/DETAILS"
'        End If
        Printer.CurrentX = LeftMargin + 100
        Printer.CurrentY = YPos
        Printer.Print PrintMatter
        
        Dim CompanyCode As String
        
        If Me.cboCompany.Text = "LEPL" Then
            PrintMatter = "Lakshya Educare Pvt. Ltd."
            CompanyCode = "LAKSH"
        Else
            PrintMatter = "MT-EDUCARE Ltd"
            CompanyCode = "MTEL"
        End If
        Printer.CurrentX = RightMargin - Printer.TextWidth(PrintMatter) - 100
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        YPos = YPos + Printer.TextHeight("X") + 100
        Printer.Line (LeftMargin, YPos)-(RightMargin, YPos)
        YMain = YPos

        YPos = YPos + 100
        Col1Left = LeftMargin + 10 * 567
        Col2Left = Col1Left + 5000
        Printer.FontBold = False
        Printer.FontSize = 9
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        PrintMatter = "Slip No:"
        Printer.CurrentX = LeftMargin + 100
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        ''''Print Slip No
        PrintMatter = Me.txtSlipNo.Text & "  [" & Me.txtCMSSlipType.Text & "]"
        Printer.CurrentX = LeftMargin + (3 * 567)
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        PrintMatter = "Slip Date:"
        Printer.CurrentX = Col1Left
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        PrintMatter = Me.txtSlipDate.Text
        Printer.CurrentX = Col1Left + (3 * 567)
        Printer.CurrentY = YPos
        Printer.Print PrintMatter
       
        YPos = YPos + Printer.TextHeight("X") + 50
        
        
        PrintMatter = "Account No:"
        Printer.CurrentX = LeftMargin + 100
        Printer.CurrentY = YPos
        Printer.Print PrintMatter
        
        PrintMatter = ""
        If Me.cboCompany.Text = "MTEX" Then
            PrintMatter = "ICICI Bank Ltd - MH-STB-ENG - 015105012577"
        ElseIf Me.cboCompany.Text = "MTMA" Then
            PrintMatter = "ICICI Bank Ltd - MH-STB-MAR - 015105012578"
        ElseIf Me.cboCompany.Text = "MTCO" Then
            PrintMatter = "ICICI Bank Ltd - MH-COM - 015105012579"
        ElseIf Me.cboCompany.Text = "MTGU" Then
            PrintMatter = "ICICI Bank Ltd - GUJ - 015105012580"
        ElseIf Me.cboCompany.Text = "MTTA" Then
            PrintMatter = "ICICI Bank Ltd - TN - 015105012581"
        ElseIf Me.cboCompany.Text = "MTRM" Then
            PrintMatter = "ICICI Bank Ltd - MH-PUNE - 015105012582"
        ElseIf Me.cboCompany.Text = "MTCB" Then
            PrintMatter = "ICICI Bank Ltd - MH-CBSE - 015105012583"
        ElseIf Me.cboCompany.Text = "MTIC" Then
            PrintMatter = "ICICI Bank Ltd - MH-ICSE - 015105012584"
        ElseIf Me.cboCompany.Text = "MTSC" Then
            PrintMatter = "ICICI Bank Ltd - MH-SCI - 015105012585"
        ElseIf Me.cboCompany.Text = "MTKA" Then
            PrintMatter = "ICICI Bank Ltd - KA - 015105012587"
        ElseIf Me.cboCompany.Text = "MTIS" Then
            PrintMatter = "ICICI Bank Ltd - INK-STB-VIII - 015105012588"
        ElseIf Me.cboCompany.Text = "LEPL" Then
            PrintMatter = "124405000213"
        End If
        
        Printer.CurrentX = LeftMargin + (3 * 567)
        Printer.CurrentY = YPos
        Printer.Print PrintMatter
        
        PrintMatter = "Company Code: " & CompanyCode
        Printer.CurrentX = Col1Left
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        PrintMatter = "Division Code: " & Me.cboCompany.Text
        Printer.CurrentX = Col1Left + (4 * 567)
        Printer.CurrentY = YPos
        Printer.Print PrintMatter
        
        YPos = YPos + Printer.TextHeight("X") + 50
        
        PrintMatter = "Pickup Point:"
        Printer.CurrentX = LeftMargin + 100
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        PrintMatter = ModInit.LocationName
        Printer.CurrentX = LeftMargin + (3 * 567)
        Printer.CurrentY = YPos
        Printer.Print PrintMatter
        
        PrintMatter = "Cheque Count:"
        Printer.CurrentX = Col1Left
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        PrintMatter = Me.txtChequeCnt.Text & "          " & "(Page " & Val(X) & " of " & Val(TPAGES) & " )"
        Printer.CurrentX = Col1Left + (3 * 567)
        Printer.CurrentY = YPos
        Printer.Print PrintMatter
        
        YPos = YPos + Printer.TextHeight("X") + 50
        
        
        Printer.Line (LeftMargin, YPos)-(RightMargin, YPos)
        YStart1 = YPos
        Printer.FontSize = 8
        'Print column headings
        Col1Left = LeftMargin + 500
        Col2Left = Col1Left + 1000
        Col3Left = Col2Left + 1000
        Col4Left = Col3Left + 1000
        Col5Left = Col4Left + 1500
        Col6Left = Col5Left + 1500

        YPos = YPos + 50

        Printer.CurrentY = YPos
        PrintMatter = "No."
        Printer.CurrentX = LeftMargin + (Col1Left - LeftMargin) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter
        
        Printer.CurrentY = YPos
        PrintMatter = "Check No."
        Printer.CurrentX = Col1Left + 100
        Printer.Print PrintMatter
        
        Printer.CurrentY = YPos
        PrintMatter = "Barcode No"
        Printer.CurrentX = Col2Left + 100
        Printer.Print PrintMatter
        
        Printer.CurrentY = YPos
        PrintMatter = "Check Date"
        Printer.CurrentX = Col3Left + 100
        Printer.Print PrintMatter
        
        Printer.CurrentY = YPos
        PrintMatter = "Drawee Bank Code"
        Printer.CurrentX = Col4Left + 100
        Printer.Print PrintMatter

        Printer.CurrentY = YPos
        PrintMatter = "Amount"
        Printer.CurrentX = Col5Left + (Col5Left - Col4Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos
        PrintMatter = "Centre"
        Printer.CurrentX = Col6Left + 100
        Printer.Print PrintMatter

        YPos = YPos + Printer.TextWidth("X") + 100

        'Horz divider (bottom of column titles)
        Printer.Line (LeftMargin, YPos)-(RightMargin, YPos)
        YPos = YPos - 200

        'Determines where to stop the loop
        If X = TPAGES Then
            ENDNUM = TROWS
        Else
            ENDNUM = PrintLines * X
        End If

        For Y = 1 + (PrintLines * (X - 1)) To ENDNUM
            'Determines where to set CurrentY Position W/ Regard to Next Sheet
            If X > 1 Then
                nxtrow = 300 * (Y - (PrintLines * (X - 1)))
            Else
                nxtrow = 300 * Y
            End If

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Val(Y)
            Printer.CurrentX = LeftMargin + (Col1Left - LeftMargin) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter
            
            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 3)
            Printer.CurrentX = Col1Left + 100 '(Col2Left - Col1Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter
            
            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 5)
            Printer.CurrentX = Col2Left + 100
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 10)    'ChkDate
            Printer.CurrentX = Col3Left + 100
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 2)
            Printer.CurrentX = Col4Left + 100
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 4)
            Printer.CurrentX = Col5Left + (Col5Left - Col4Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter
        
            SlipAmtTotal = SlipAmtTotal + Val(Grid.TextMatrix(Y, 4))
            
            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 8)
            Printer.CurrentX = Col6Left + 100
            Printer.Print PrintMatter

            If Y <> ENDNUM Then
                Printer.Line (LeftMargin, YPos + nxtrow + Printer.TextHeight("X") + 50)-(RightMargin, YPos + nxtrow + Printer.TextHeight("X") + 50)
            End If
        Next Y

        YPos = YPos + nxtrow + Printer.TextHeight("X") + 50
        Printer.Line (LeftMargin, YPos)-(RightMargin, YPos)

        Printer.Line (Col1Left, YStart1)-(Col1Left, YPos)
        Printer.Line (Col2Left, YStart1)-(Col2Left, YPos)
        Printer.Line (Col3Left, YStart1)-(Col3Left, YPos)
        Printer.Line (Col4Left, YStart1)-(Col4Left, YPos)
        Printer.Line (Col5Left, YStart1)-(Col5Left, YPos)
        Printer.Line (Col6Left, YStart1)-(Col6Left, YPos)
        Printer.Line (LeftMargin, YStart)-(LeftMargin, YPos)
        Printer.Line (RightMargin, YStart)-(RightMargin, YPos)

        YPos = YPos + 100

        If X = TPAGES Then

            Printer.CurrentY = YPos
            PrintMatter = ""
            Printer.CurrentX = LeftMargin + 100
            Printer.Print PrintMatter

            Printer.CurrentY = YPos
            PrintMatter = "TOTAL ->"
            Printer.CurrentX = Col5Left - Printer.TextWidth(PrintMatter) - 150
            Printer.Print PrintMatter

            Printer.CurrentY = YPos
            PrintMatter = Format(SlipAmtTotal, "#,##,##,##0.00")
            Printer.CurrentX = Col5Left + (Col5Left - Col4Left) / 2 - (Printer.TextWidth(PrintMatter) / 2) ' RightMargin - Printer.TextWidth(printmatter) - 150
            Printer.Print PrintMatter

            YPos = YPos + Printer.TextWidth("X") + 100

            Printer.Line (Col4Left, YStart1)-(Col4Left, YPos)
            Printer.Line (Col5Left, YStart1)-(Col5Left, YPos)
'            Printer.Line (Col6Left, YStart1 + Printer.TextWidth("X") + 150)-(Col6Left, YPos)
            Printer.Line (LeftMargin, YPos)-(RightMargin, YPos)

            YPos = YPos + Printer.TextWidth("X") + 100
'            Printer.CurrentY = YPos
'            PrintMatter = "1. White  : Citibank Copy"
'            Printer.CurrentX = LeftMargin + 150
'            Printer.Print PrintMatter

            YPos = YPos + Printer.TextWidth("X") + 100
'            Printer.CurrentY = YPos
'            PrintMatter = "2. Blue   : Co-Ordinator Copy"
'            Printer.CurrentX = LeftMargin + 150
'            Printer.Print PrintMatter

            YPos = YPos + Printer.TextWidth("X") + 100
'            Printer.CurrentY = YPos
'            PrintMatter = "3. Yellow : Customer Copy"
'            Printer.CurrentX = LeftMargin + 150
'            Printer.Print PrintMatter

            YPos = YPos + Printer.TextWidth("X") + 100
            Printer.CurrentY = YPos
            PrintMatter = "Customer's Signature  (Date & Time)"
            Printer.CurrentX = RightMargin - Printer.TextWidth(PrintMatter) - 2000
            Printer.Print PrintMatter

            XPos = RightMargin - Printer.TextWidth(PrintMatter) - 2000
            Printer.Line (XPos, YPos - 100)-(XPos + Printer.TextWidth(PrintMatter) + 50, YPos - 100)


            YPos = YPos + Printer.TextHeight("X") + 50
            Printer.Line (LeftMargin, YPos)-(RightMargin, YPos)
            Printer.Line (LeftMargin, YStart)-(LeftMargin, YPos)
            Printer.Line (RightMargin, YStart)-(RightMargin, YPos)
        End If

        Printer.FontSize = 7
        YPos = BottomMargin - Printer.TextHeight("X") - 300

'        If OptLocal.Value = True Then
'            If intCopies = 4 Then
'                printmatter = "Customer Copy"
'            ElseIf intCopies = 3 Then
'                printmatter = "Co-Ordinator Copy"
'            Else
'                printmatter = "Citibank Copy"
'            End If
'        Else
'            If intCopies = 3 Then
'                PrintMatter = "Customer Copy"
'            ElseIf intCopies = 2 Then
'                PrintMatter = "Co-Ordinator Copy"
'            Else
'                PrintMatter = "Citibank Copy"
'            End If
'        End If

'        PrintMatter1 = "(Page of " & Val(X) & " of " & Val(TPAGES) & " )"
'
'        Printer.CurrentY = YPos
'        Printer.CurrentX = LeftMargin + ((RightMargin - LeftMargin) / 2) - (Printer.TextWidth(printmatter) / 2)
'        Printer.Print printmatter

        YPos = YPos + Printer.TextHeight("X") + 50
        Printer.CurrentY = YPos
        Printer.CurrentX = LeftMargin + ((RightMargin - LeftMargin) / 2) - (Printer.TextWidth(PrintMatter1) / 2)
        Printer.Print PrintMatter1
        Printer.NewPage
    Next X
Next intCopies
Printer.EndDoc
Printer.Orientation = cdlPortrait
Exit Sub

ErrHandler:

'Printer.Orientation = cdlPortrait
Printer.EndDoc
End Sub

Private Sub cmdSMS_Click()
On Error Resume Next
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

Dim MobileNo As String
Dim MsgString As String

Me.MousePointer = vbHourglass
lblSMSCount.Visible = True
lblSMSCount.Caption = "0"

rs1.Open "Select * from ASPDC_DispatchSlipDetails where CMSDoneFlag =1 and CMSSlipNo ='" & Me.txtSlipNo.Text & "' and Location_Code ='" & ModInit.LocationCode & "' and (SMS_Flag =0 or SMS_Flag is Null) AND cms_VERIFyFlag =1", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        'Find Mobile number of student
        If "" & rs1!CMSSBEntryCode <> "" Then
'            rs2.Open "select isnull(Pmobileno,'') as Pmobileno from studentbatch sb inner join student s on sb.yearname = s.yearname and sb.institutecode = s.institutecode and sb.liccode = s.liccode and sb.studentcode = s.studentcode where sbentrycode ='" & rs1!CMSSBEntryCode & "'", cn1, adOpenDynamic, adLockReadOnly
'            If Not (rs2.BOF And rs2.EOF) Then
                If "" & rs1!SMS_MobileNo <> "" Then
                    MobileNo = rs1!SMS_MobileNo
                    
                    MsgString = "Dear Parent, Please note that your cheque amounting to Rs. " & Format(rs1!CCChequeAmt, "0.00") & " will be deposited on " & Format(rs1!CMSSlipDate, "dd Mmm yyyy") & ". Regards, MT Educare Ltd"
                    
                    'MsgString = "Dear Parent, Due to a system change at our end, cheques issued by you dated 15th Jan'15 and after, would be deposited on 12th Feb'15. Request you to maintain sufficient balance, to honour the cheque. Thanks, MT Educare"
                    
                    'Send Message
                    ModInit.WaitFlag = True
                    FrmMsgSender.SendSMS MobileNo, MsgString, rs1!CCChqIdNo
                    Do While ModInit.WaitFlag = True
                        DoEvents
                    Loop
                    
                    lblSMSCount.Caption = Val(lblSMSCount.Caption) + 1
                    DoEvents
                Else
                    UpdateOutput 3, rs1!CCChqIdNo
                End If
'            End If
'            rs2.Close
        End If
        
        rs1.MoveNext
    Loop
End If
rs1.Close
cn1.Close
MsgBox "Done.", vbInformation + vbOKOnly

Me.MousePointer = 0
End Sub


Private Sub UpdateOutput(SuccessFlag As Integer, Barcode As String)
On Error Resume Next
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

rs1.Open "Select * from ASPDC_DispatchSlipDetails where CCChqIdNo ='" & Barcode & "'", cn1, adOpenDynamic, adLockPessimistic
If Not (rs1.BOF And rs1.EOF) Then
    rs1!SMS_Flag = 3    'Mobile No not found
    rs1!SMS_MobileNo = "Not Present"
    rs1!SMS_Date = Date
    rs1.Update
End If
rs1.Close
cn1.Close

End Sub

Private Sub cmdVerify_Click()
On Error Resume Next
Dim NewRowNo As Integer
With FrmCMSDeposit_New_Sub2
    .txtSlipNo.Text = Me.txtSlipNo.Text
    .txtSlipDate.Text = Me.txtSlipDate.Text
    .txtChequeCnt.Text = Me.txtChequeCnt.Text
    .Grid.Rows = 1
    NewRowNo = 1
    For Cnt = 1 To Grid.Rows - 1
        If Grid.TextMatrix(Cnt, 11) = "" Then
            .Grid.Rows = .Grid.Rows + 1
            .Grid.TextMatrix(NewRowNo, 1) = Grid.TextMatrix(Cnt, 3)
            .Grid.TextMatrix(NewRowNo, 2) = Grid.TextMatrix(Cnt, 4)
            .Grid.TextMatrix(NewRowNo, 3) = Grid.TextMatrix(Cnt, 5)
            .Grid.TextMatrix(NewRowNo, 4) = Grid.TextMatrix(Cnt, 6)
            .Grid.TextMatrix(NewRowNo, 5) = Grid.TextMatrix(Cnt, 7)
            .Grid.TextMatrix(NewRowNo, 6) = Grid.TextMatrix(Cnt, 8)
            .Grid.TextMatrix(NewRowNo, 7) = Grid.TextMatrix(Cnt, 9)
            .Grid.TextMatrix(NewRowNo, 8) = Grid.TextMatrix(Cnt, 10)
            .Grid.TextMatrix(NewRowNo, 0) = Grid.TextMatrix(Cnt, 11)
            NewRowNo = NewRowNo + 1
        Else
            'MsgBox Grid.TextMatrix(Cnt, 11)
        End If
    Next
    .txtChequeCnt.Text = .Grid.Rows - 1
    .Show vbModal
End With
End Sub

Private Sub Form_Load()
On Error Resume Next
txtFlag.Text = "ADD"
TxtUserName.Text = ModInit.PDCUserName
With Grid
    '<Deposit Flag|<Instrument No.|<Instrument Date|>Instrument Amount|<Bank Name|<Name of the Student|<Course|<Form No.|<Academic Year|<RcptCode|<SBEntryCode|<ChequeIdNo
    .ColWidth(0) = (.Width - 350) / (.Cols - 1)
    .ColWidth(1) = 0
    .ColWidth(2) = .ColWidth(0)
    .ColWidth(3) = .ColWidth(0)
    .ColWidth(4) = .ColWidth(0)
    .ColWidth(5) = .ColWidth(0)
    .ColWidth(6) = .ColWidth(0)
    .ColWidth(7) = .ColWidth(0)
    .ColWidth(8) = .ColWidth(0)
    .Rows = 1
End With

End Sub

Public Sub FillCompany(SlipCode As String, FillOption As Integer)
On Error Resume Next
txtCMSSlipTypeCode.Text = FillOption
Me.txtSlipNo.Text = SlipCode
Me.txtSlipDate.Text = Format(FrmCMSDeposit_New.dtSlip.Value, "dd Mm yyyy")
    
If FillOption = 1 Then
    Me.txtCMSSlipType.Text = "ICICI Bank"
ElseIf FillOption = 2 Then
    Me.txtCMSSlipType.Text = "Local Non-ICICI Bank"
ElseIf FillOption = 3 Then
    Me.txtCMSSlipType.Text = "Outstation Bank"
Else
    Me.txtCMSSlipType.Text = "Non MICR Cheques"
End If

If SlipCode = "" Then Exit Sub
    
'FillOption = 1 means ICICI Bank
'FillOption = 2 means Local Non-ICICI Bank
'FillOption = 3 means Outstation
'FillOption = 4 means Non MICR

Me.MousePointer = vbHourglass
Me.cboCompany.Clear

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

'Start mapping the two entries
Dim str As String
If FillOption = 4 Then
    str = "select distinct C8.CMS_Company_Code from ASPDC_DispatchSlipDetails ASDD inner join ASPDC_DispatchSlip ASD on ASD.DispatchSlipCode = ASDD.DispatchSlipCode and ASD.Location_Code = ASDD.Location_Code inner join c008_centers C8 on C8.source_center_code = ASD.new_InstituteCode + ASD.New_LicCode where ASDD.Location_Code ='" & ModInit.LocationCode & "' and ASDD.cmsslipno ='" & SlipCode & "' and (len(isnull(ASDD.MICRNumber,''))) <> 9"
ElseIf FillOption = 3 Then
    str = "select distinct C8.CMS_Company_Code from ASPDC_DispatchSlipDetails ASDD inner join ASPDC_DispatchSlip ASD on ASD.DispatchSlipCode = ASDD.DispatchSlipCode and ASD.Location_Code = ASDD.Location_Code inner join c008_centers C8 on C8.source_center_code = ASD.new_InstituteCode + ASD.New_LicCode where ASDD.Location_Code ='" & ModInit.LocationCode & "' and ASDD.cmsslipno ='" & SlipCode & "' and right(left( MICRNumber,6),3) <> '229' and left(MICRNumber,3) <> '" & ModInit.MICRLocationCode & "'"
ElseIf FillOption = 2 Then
    str = "select distinct C8.CMS_Company_Code from ASPDC_DispatchSlipDetails ASDD inner join ASPDC_DispatchSlip ASD on ASD.DispatchSlipCode = ASDD.DispatchSlipCode and ASD.Location_Code = ASDD.Location_Code inner join c008_centers C8 on C8.source_center_code = ASD.new_InstituteCode + ASD.New_LicCode where ASDD.Location_Code ='" & ModInit.LocationCode & "' and ASDD.cmsslipno ='" & SlipCode & "' and right(left( MICRNumber,6),3) <> '229' and left(MICRNumber,3) ='" & ModInit.MICRLocationCode & "'"
ElseIf FillOption = 1 Then
    str = "select distinct C8.CMS_Company_Code from ASPDC_DispatchSlipDetails ASDD inner join ASPDC_DispatchSlip ASD on ASD.DispatchSlipCode = ASDD.DispatchSlipCode and ASD.Location_Code = ASDD.Location_Code inner join c008_centers C8 on C8.source_center_code = ASD.new_InstituteCode + ASD.New_LicCode where ASDD.Location_Code ='" & ModInit.LocationCode & "' and ASDD.cmsslipno ='" & SlipCode & "' and right(left( MICRNumber,6),3) = '229'"
End If
'str = "Select * from ASPDC_DispatchSlipDetails where DispatchSlipCode ='" & Trim(UCase(SlipCode)) & "'"
rs1.Open str, cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        
        Me.cboCompany.AddItem rs1!CMS_Company_Code
        rs1.MoveNext
    Loop
    cboCompany.ListIndex = 0
End If
rs1.Close
cn1.Close

Me.MousePointer = 0
End Sub


Public Sub FillGrid(SlipCode As String, FillOption As Integer)
On Error Resume Next
Me.txtChequeCnt.Text = "0"
Me.txtICICIDSNo.Text = ""

If SlipCode = "" Then Exit Sub

'FillOption = 1 means ICICI Bank
'FillOption = 2 means Local Non-ICICI Bank
'FillOption = 3 means Outstation
'FillOption = 4 means Non MICR

Me.MousePointer = vbHourglass
Grid.Rows = 1

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

'Start mapping the two entries
Dim str As String
If FillOption = 4 Then
    str = "select ASDD.CCChequeNo, ASDD.CCChequeAmt, ASDD.CCChqIDNo, ASDD.MICRNumber, ASDD.ICICIBankDSNo, ASDD.ICICIBankDSChequeEntryNo, ASDD.TranCode, ASD.MISInstituteCode, ASD.LicCode, C8.CMS_Company_Code, C8.Target_Center_Name, ASDD.CMSSBEntryCode, ASDD.CCChequeDate, ASDD.CMS_VerifyFlag from ASPDC_DispatchSlipDetails ASDD inner join ASPDC_DispatchSlip ASD on ASD.DispatchSlipCode = ASDD.DispatchSlipCode and ASD.Location_Code = ASDD.Location_Code inner join c008_centers C8 on C8.source_center_code = ASD.new_InstituteCode + ASD.new_LicCode where ASDD.Location_Code ='" & ModInit.LocationCode & "' and ASDD.cmsslipno ='" & SlipCode & "' and (len(isnull(ASDD.MICRNumber,''))) <> 9 and C8.CMS_Company_Code ='" & Me.cboCompany.Text & "' order by ASDD.ICICIBankDSNo, MISInstituteCode, C8.Target_Center_Name"
ElseIf FillOption = 3 Then
    str = "select ASDD.CCChequeNo, ASDD.CCChequeAmt, ASDD.CCChqIDNo, ASDD.MICRNumber, ASDD.ICICIBankDSNo, ASDD.ICICIBankDSChequeEntryNo, ASDD.TranCode, ASD.MISInstituteCode, ASD.LicCode, C8.CMS_Company_Code, C8.Target_Center_Name, ASDD.CMSSBEntryCode, ASDD.CCChequeDate, ASDD.CMS_VerifyFlag from ASPDC_DispatchSlipDetails ASDD inner join ASPDC_DispatchSlip ASD on ASD.DispatchSlipCode = ASDD.DispatchSlipCode and ASD.Location_Code = ASDD.Location_Code inner join c008_centers C8 on C8.source_center_code = ASD.new_InstituteCode + ASD.new_LicCode where ASDD.Location_Code ='" & ModInit.LocationCode & "' and ASDD.cmsslipno ='" & SlipCode & "' and right(left( MICRNumber,6),3) <> '229' and left(MICRNumber,3) <> '" & ModInit.MICRLocationCode & "' and C8.CMS_Company_Code ='" & Me.cboCompany.Text & "' order by ASDD.ICICIBankDSNo, MISInstituteCode, C8.Target_Center_Name"
ElseIf FillOption = 2 Then
    str = "select ASDD.CCChequeNo, ASDD.CCChequeAmt, ASDD.CCChqIDNo, ASDD.MICRNumber, ASDD.ICICIBankDSNo, ASDD.ICICIBankDSChequeEntryNo, ASDD.TranCode, ASD.MISInstituteCode, ASD.LicCode, C8.CMS_Company_Code, C8.Target_Center_Name, ASDD.CMSSBEntryCode, ASDD.CCChequeDate, ASDD.CMS_VerifyFlag from ASPDC_DispatchSlipDetails ASDD inner join ASPDC_DispatchSlip ASD on ASD.DispatchSlipCode = ASDD.DispatchSlipCode and ASD.Location_Code = ASDD.Location_Code inner join c008_centers C8 on C8.source_center_code = ASD.new_InstituteCode + ASD.new_LicCode where ASDD.Location_Code ='" & ModInit.LocationCode & "' and ASDD.cmsslipno ='" & SlipCode & "' and right(left( MICRNumber,6),3) <> '229' and left(MICRNumber,3) ='" & ModInit.MICRLocationCode & "' and C8.CMS_Company_Code ='" & Me.cboCompany.Text & "' order by ASDD.ICICIBankDSNo, MISInstituteCode, C8.Target_Center_Name"
ElseIf FillOption = 1 Then
    str = "select ASDD.CCChequeNo, ASDD.CCChequeAmt, ASDD.CCChqIDNo, ASDD.MICRNumber, ASDD.ICICIBankDSNo, ASDD.ICICIBankDSChequeEntryNo, ASDD.TranCode, ASD.MISInstituteCode, ASD.LicCode, C8.CMS_Company_Code, C8.Target_Center_Name, ASDD.CMSSBEntryCode, ASDD.CCChequeDate, ASDD.CMS_VerifyFlag from ASPDC_DispatchSlipDetails ASDD inner join ASPDC_DispatchSlip ASD on ASD.DispatchSlipCode = ASDD.DispatchSlipCode and ASD.Location_Code = ASDD.Location_Code inner join c008_centers C8 on C8.source_center_code = ASD.new_InstituteCode + ASD.new_LicCode where ASDD.Location_Code ='" & ModInit.LocationCode & "' and ASDD.cmsslipno ='" & SlipCode & "' and right(left( MICRNumber,6),3) = '229' and C8.CMS_Company_Code ='" & Me.cboCompany.Text & "' order by ASDD.ICICIBankDSNo, MISInstituteCode, C8.Target_Center_Name"
End If

'str = "Select * from ASPDC_DispatchSlipDetails where DispatchSlipCode ='" & Trim(UCase(SlipCode)) & "'"
rs1.Open str, cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        Grid.Rows = Grid.Rows + 1
        Cnt = Grid.Rows - 1
        '<Cheque Date|<Verified Status|<ICICI Bank DS No|<Entry No|<Tran Code
        
        Grid.TextMatrix(Cnt, 0) = Cnt
        Grid.TextMatrix(Cnt, 1) = "Student Name"
        Grid.TextMatrix(Cnt, 2) = Mid(rs1!MICRNumber, 4, 3)
        Grid.TextMatrix(Cnt, 3) = rs1!CCChequeNo
        Grid.TextMatrix(Cnt, 4) = Format(rs1!CCChequeAmt, "0.00")
        Grid.TextMatrix(Cnt, 5) = rs1!CCChqIdNo
        Grid.TextMatrix(Cnt, 6) = rs1!MICRNumber
        Grid.TextMatrix(Cnt, 7) = rs1!MISInstituteCode
        Grid.TextMatrix(Cnt, 8) = rs1!LicCode & "-" & rs1!Target_Center_Name
        Grid.TextMatrix(Cnt, 9) = rs1!CMSSBEntryCode
        Grid.TextMatrix(Cnt, 10) = Format(rs1!CCChequeDate, "dd Mm YYYY")
        If "" & rs1!CMS_VerifyFlag = "1" Then
            Grid.TextMatrix(Cnt, 11) = "Found"
        Else
            Grid.TextMatrix(Cnt, 11) = ""
        End If
        Grid.TextMatrix(Cnt, 12) = "" & rs1!ICICIBankDSNo
        Grid.TextMatrix(Cnt, 13) = "" & rs1!ICICIBankDSChequeEntryNo
        If "" & rs1!TranCode = "" Then
            Grid.TextMatrix(Cnt, 14) = "10"
        Else
            Grid.TextMatrix(Cnt, 14) = "" & rs1!TranCode
        End If
'        Grid.TextMatrix(Cnt, 1) = rs1!CCChequeDate
        
        If txtICICIDSNo.Text = "" And "" & rs1!ICICIBankDSNo <> "" Then
            txtICICIDSNo.Text = "" & rs1!ICICIBankDSNo
        End If
    
        rs1.MoveNext
    Loop
End If
rs1.Close
cn1.Close
Me.txtChequeCnt.Text = Grid.Rows - 1
Me.MousePointer = 0
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

If ModInit.PDCUserName = "Mithun" Then
    FrmCMSDeposit_New_Sub4.Show vbModal
End If

End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
On Error Resume Next
MsgBox KeyAscii
If Grid.Rows = 1 Then Exit Sub

'If KeyAscii = 114 Or KeyAscii = 82 Then
'    'Remove entry from cms
'    Dim cn1 As ADODB.Connection
'    Set cn1 = New ADODB.Connection
'
'    cn1.ConnectionString = ModInit.ConnectStringOnline
'    cn1.Open
'
'    Dim rs1 As ADODB.Recordset
'    Set rs1 = New ADODB.Recordset
'
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordse
'
'    'Delete entry from T005_CMS_Data
'    rs1.Open "Delete from T005_CMS_Data where Pay_insnum ='" & Me.txtSlipNo.Text & "' and Cur_SB_Code ='" & Grid.TextMatrix(Grid.RowSel, 6) & "' and Pay_InsNum ='" & Grid.TextMatrix(Grid.RowSel, 2) & "'", cn1, adOpenDynamic, adLockPessimistic
'
'    'Change CMS Status of the cheque
'    rs1.Open
'
'End If

End Sub


Private Sub mnuCTSUploadFile_Click()
'On Error Resume Next
'Dim res As Integer
'res = MsgBox("You are about to generate CTS File for ICICI Bank.  This will take some time.  Do you want to proceed?", vbYesNo + vbQuestion)
'If res = 7 Then Exit Sub
'
''Add header row
'Dim DSRowNo As Integer
'
'With Me.BankExportGrid
'    .Rows = 0
'    For CompCnt = 0 To Me.cboCompany.ListCount - 1
'        cboCompany.ListIndex = CompCnt
'        DoEvents
'
'
'
'        .Rows = .Rows + 1
'        DSRowNo = .Rows - 1
'
'        .TextMatrix(.Rows - 1, 0) = "DS"
'        If Me.cboCompany.Text = "LEPL" Then
'            .TextMatrix(.Rows - 1, 1) = "LAKSH"
'        Else
'            .TextMatrix(.Rows - 1, 1) = "MTEL"
'        End If
'        .TextMatrix(.Rows - 1, 2) = "LCCBRN"
'        .TextMatrix(.Rows - 1, 3) = ""
'        .TextMatrix(.Rows - 1, 4) = ""
'        .TextMatrix(.Rows - 1, 5) = Me.cboCompany.Text
'        If Me.txtICICIDSNo.Text <> "" Then
'            .TextMatrix(.Rows - 1, 6) = "'" & Me.txtICICIDSNo.Text
'        Else
'            .TextMatrix(.Rows - 1, 6) = Me.txtSlipNo.Text & "-" & Me.txtCMSSlipType.Text
'        End If
'        .TextMatrix(.Rows - 1, 7) = "'" & Me.txtSlipDate.Text
'        .TextMatrix(.Rows - 1, 8) = Me.txtChequeCnt.Text
'        .TextMatrix(.Rows - 1, 9) = "0"
'        .TextMatrix(.Rows - 1, 10) = ""
'        .TextMatrix(.Rows - 1, 11) = ""
'        .TextMatrix(.Rows - 1, 12) = ""
'        .TextMatrix(.Rows - 1, 13) = ""
'        .TextMatrix(.Rows - 1, 14) = ""
'        .TextMatrix(.Rows - 1, 15) = ModInit.LocationName
'
'        'Start rows for DSI entry
'        Dim SlipAmount As Double
'        SlipAmount = 0
'
'        For Cnt = 1 To Grid.Rows - 1
'            .Rows = .Rows + 1
'            .TextMatrix(.Rows - 1, 0) = "DSI"
'            .TextMatrix(.Rows - 1, 1) = Grid.TextMatrix(Cnt, 3)             'INSTRUMENT NO
'            .TextMatrix(.Rows - 1, 2) = "'" & Replace(Grid.TextMatrix(Cnt, 10), " ", "")          'INSTRUMENT DATE
'            .TextMatrix(.Rows - 1, 3) = "'" & Grid.TextMatrix(Cnt, 14)             'INSTRUMENT TYPE
'            .TextMatrix(.Rows - 1, 4) = ""
'            .TextMatrix(.Rows - 1, 5) = "'" & Grid.TextMatrix(Cnt, 2)             'DRAWN ON BANK
'            .TextMatrix(.Rows - 1, 6) = "'" & Right(Grid.TextMatrix(Cnt, 6), 3)   'DRAWN ON BRANCH
'            .TextMatrix(.Rows - 1, 7) = Grid.TextMatrix(Cnt, 4)             'INSTRUMENT AMOUNT
'            .TextMatrix(.Rows - 1, 8) = ""
'            .TextMatrix(.Rows - 1, 9) = "NAME NOT MENTION"
'            .TextMatrix(.Rows - 1, 10) = "'" & Grid.TextMatrix(Cnt, 9)      'SBEntrycode
'            .TextMatrix(.Rows - 1, 11) = Me.txtSlipNo.Text                                  'CMS Number
'            .TextMatrix(.Rows - 1, 12) = "'" & Grid.TextMatrix(Cnt, 7) & Left(Grid.TextMatrix(Cnt, 7), 2) 'Centre Code
'            .TextMatrix(.Rows - 1, 13) = Grid.TextMatrix(Cnt, 6)
'
'            SlipAmount = SlipAmount + Val(Grid.TextMatrix(Cnt, 4))
'
'            DoEvents
'        Next
'
'        .TextMatrix(DSRowNo, 9) = SlipAmount
'
'
'    Next
'End With
'
'GenerateCTSFile False

End Sub

Private Sub mnuProfundUploadFile_Click()
On Error Resume Next
If txtCMSSlipTypeCode.Text = "4" Then
    MsgBox "Export for Bank option is not allowed for this type.", vbCritical + vbOKOnly
    Exit Sub
End If

Dim res As Integer
res = MsgBox("You are about to generate Profound Upload File for Kotak Bank.  This will take some time.  Do you want to proceed?", vbYesNo + vbQuestion)
If res = 7 Then Exit Sub

'Save file
Dim FSO As FileSystemObject
Set FSO = New FileSystemObject

Dim UploadFolder As String
UploadFolder = App.Path & "\UploadFiles\" & Me.txtSlipDate.Text & "\"

If FSO.FolderExists(UploadFolder) = False Then
    'Create folder
    FSO.CreateFolder UploadFolder
End If


'Add header row
Dim DSRowNo As Integer
On Error Resume Next
Dim OutputString As String
Dim wrkErrorMessage As String
Dim wrkOutputFile As String
Dim wrkProjectName As String
Dim ColCnt, RowCnt As Integer
Dim SlipAmount As Double
Dim ChequeCnt As Integer

Dim CurDSNo As String
CurDSNo = "-"
With Me.BankExportGrid
    For CompCnt = 0 To Me.cboCompany.ListCount - 1
        cboCompany.ListIndex = CompCnt
        DoEvents
        
        For Cnt = 1 To Grid.Rows - 1
            If Grid.TextMatrix(Cnt, 12) <> CurDSNo Then
                'First export existing data
                If .Rows > 1 Then
                    
                    .TextMatrix(DSRowNo, 9) = SlipAmount
                    .TextMatrix(DSRowNo, 8) = ChequeCnt
                    OutputString = ""
                    
                    wrkProjectName = "Upload File - " & CurDSNo & " - " & Me.txtSlipDate.Text
                    wrkOutputFile = UploadFolder & "" & Me.txtCMSSlipType.Text & "-" & CurDSNo & ".xls"   'cmd.FileName
                    

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

                    
                    '  Save the Grid as a File
                    If (SaveExcelWorksheet_Profund(wrkProjectName, wrkOutputFile, wrkErrorMessage) = False) Then
                        MsgBox "Error in data transfer."
                        Exit Sub
                    End If
                    
                    GenerateCTSFile False, CurDSNo

                
                End If
                
                'Start loading new data
                CurDSNo = Grid.TextMatrix(Cnt, 12)
                .Rows = 0
                
                '****Start New part for KOTAK
                'DS Header
                .Rows = .Rows + 1
                DSRowNo = .Rows - 1
                
                .TextMatrix(.Rows - 1, 0) = "DS"
                .TextMatrix(.Rows - 1, 1) = "CUST CODE"
                .TextMatrix(.Rows - 1, 2) = "PRO CODE"
                .TextMatrix(.Rows - 1, 3) = "SUB CUST CODE"
                .TextMatrix(.Rows - 1, 4) = "HIER CODE"
                .TextMatrix(.Rows - 1, 5) = "DIV CODE"
                .TextMatrix(.Rows - 1, 6) = "DS. NO."
                .TextMatrix(.Rows - 1, 7) = "DATE"
                .TextMatrix(.Rows - 1, 8) = "TOTAL COUNT"
                .TextMatrix(.Rows - 1, 9) = "TOTAL AMT"
                .TextMatrix(.Rows - 1, 10) = ""
                .TextMatrix(.Rows - 1, 11) = ""
                .TextMatrix(.Rows - 1, 12) = ""
                .TextMatrix(.Rows - 1, 13) = ""
                .TextMatrix(.Rows - 1, 14) = ""
                .TextMatrix(.Rows - 1, 15) = ""
                
                '****End new part for KOTAK
                
        
                .Rows = .Rows + 1
                DSRowNo = .Rows - 1
                
                .TextMatrix(.Rows - 1, 0) = "DS"
                If Me.cboCompany.Text = "LEPL" Then
                    .TextMatrix(.Rows - 1, 1) = "LAKSH"
                    .TextMatrix(.Rows - 1, 5) = "  "
                Else
                    .TextMatrix(.Rows - 1, 1) = Me.cboCompany.Text
                    .TextMatrix(.Rows - 1, 5) = "  "
                End If
                .TextMatrix(.Rows - 1, 2) = "LCCBRN"
                .TextMatrix(.Rows - 1, 3) = ""
                .TextMatrix(.Rows - 1, 4) = ""
                
                If CurDSNo <> "" Then
                    .TextMatrix(.Rows - 1, 6) = "'" & CurDSNo 'Me.txtICICIDSNo.Text
                Else
                    .TextMatrix(.Rows - 1, 6) = Me.txtSlipNo.Text & "-" & Me.txtCMSSlipType.Text
                End If
                .TextMatrix(.Rows - 1, 7) = "'" & Replace(Me.txtSlipDate.Text, " ", "")
                .TextMatrix(.Rows - 1, 8) = ChequeCnt
                .TextMatrix(.Rows - 1, 9) = "0"
                .TextMatrix(.Rows - 1, 10) = ""
                .TextMatrix(.Rows - 1, 11) = ""
                .TextMatrix(.Rows - 1, 12) = ""
                .TextMatrix(.Rows - 1, 13) = ""
                .TextMatrix(.Rows - 1, 14) = ""
                .TextMatrix(.Rows - 1, 15) = "" ' ModInit.LocationName
                
                
                '****Start New part for KOTAK
                'DSI Header
                .Rows = .Rows + 1
                
                
                .TextMatrix(.Rows - 1, 0) = "DSI"
                .TextMatrix(.Rows - 1, 1) = "CHQ NO."
                .TextMatrix(.Rows - 1, 2) = "CHQ DATE"
                .TextMatrix(.Rows - 1, 3) = "TRAN CODE"
                .TextMatrix(.Rows - 1, 4) = " "
                .TextMatrix(.Rows - 1, 5) = "BANK CODE"
                .TextMatrix(.Rows - 1, 6) = "BR CODE"
                .TextMatrix(.Rows - 1, 7) = "AMT"
                .TextMatrix(.Rows - 1, 8) = "DRAWER CODE"
                .TextMatrix(.Rows - 1, 9) = "DRAWER NAME"
                .TextMatrix(.Rows - 1, 10) = "SB ENTRY CODE"
                .TextMatrix(.Rows - 1, 11) = "CMS DATE"
                .TextMatrix(.Rows - 1, 12) = "CENTER CODE"
                .TextMatrix(.Rows - 1, 13) = "MICR NO"
                .TextMatrix(.Rows - 1, 14) = ""
                .TextMatrix(.Rows - 1, 15) = ""
                
                '****End new part for KOTAK
                
                
                'Start rows for DSI entry
                
                SlipAmount = 0
                ChequeCnt = 0
            End If
        
        
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = "DSI"
            .TextMatrix(.Rows - 1, 1) = Grid.TextMatrix(Cnt, 3)             'INSTRUMENT NO
            .TextMatrix(.Rows - 1, 2) = "'" & Replace(Grid.TextMatrix(Cnt, 10), " ", "")          'INSTRUMENT DATE
            .TextMatrix(.Rows - 1, 3) = "'" & Grid.TextMatrix(Cnt, 14)             'INSTRUMENT TYPE
            .TextMatrix(.Rows - 1, 4) = ""
            .TextMatrix(.Rows - 1, 5) = "'" & Grid.TextMatrix(Cnt, 2)             'DRAWN ON BANK
            .TextMatrix(.Rows - 1, 6) = "'" & Right(Grid.TextMatrix(Cnt, 6), 3)   'DRAWN ON BRANCH
            .TextMatrix(.Rows - 1, 7) = Grid.TextMatrix(Cnt, 4)             'INSTRUMENT AMOUNT
            .TextMatrix(.Rows - 1, 8) = ""
            .TextMatrix(.Rows - 1, 9) = "NAME NOT MENTION"
            If Grid.TextMatrix(Cnt, 9) <> "" Then
                .TextMatrix(.Rows - 1, 10) = "'" & Grid.TextMatrix(Cnt, 9)      'SBEntrycode
            Else
                .TextMatrix(.Rows - 1, 10) = "'0"
            End If
            .TextMatrix(.Rows - 1, 11) = Me.txtSlipNo.Text
            .TextMatrix(.Rows - 1, 12) = "'" & Grid.TextMatrix(Cnt, 7) & Left(Grid.TextMatrix(Cnt, 7), 2) 'Centre Code
            .TextMatrix(.Rows - 1, 13) = Grid.TextMatrix(Cnt, 6)
            '.TextMatrix(.Rows - 1, 14) = Grid.TextMatrix(Cnt, 5)
            SlipAmount = SlipAmount + Val(Grid.TextMatrix(Cnt, 4))
            ChequeCnt = ChequeCnt + 1
            DoEvents
        Next
        
        .TextMatrix(DSRowNo, 9) = SlipAmount
        .TextMatrix(DSRowNo, 8) = ChequeCnt
        
        OutputString = ""
        wrkProjectName = "Upload File - " & CurDSNo & " - " & Me.txtSlipDate.Text
        wrkOutputFile = UploadFolder & "" & Me.txtCMSSlipType.Text & "-" & CurDSNo & ".xls"   'cmd.FileName
        

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
        '  Save the Grid as a File
        If (SaveExcelWorksheet_Profund(wrkProjectName, wrkOutputFile, wrkErrorMessage) = False) Then
            MsgBox "Error in data transfer."
            Exit Sub
        End If
        
        GenerateCTSFile False, CurDSNo

        .Rows = 0
        
    Next
End With
MsgBox "Activity finished successfully.", vbInformation + vbOKOnly

End Sub

Private Sub GenerateCTSFile(FillBankGridFlag As Boolean, DSNo As String)
'If fillbankgridflag = false then dont fill bank grid again.  use data from it
If FillBankGridFlag = True Then


End If

'Save file
Dim FSO As FileSystemObject
Set FSO = New FileSystemObject

Dim UploadFolder As String
UploadFolder = App.Path & "\UploadFiles\" & Me.txtSlipDate.Text & "\"

If FSO.FolderExists(UploadFolder) = False Then
    'Create folder
    FSO.CreateFolder UploadFolder
End If

Dim CompName As String

Dim Cnt As Integer
Dim OutputStr As String
Dim RecCnt As Integer
RecCnt = 0
Dim RunNo As Integer
RunNo = 1
For Cnt = 0 To BankExportGrid.Rows - 1
    If BankExportGrid.TextMatrix(Cnt, 0) = "DS" Then
        'Header Entry
        AccountNo = "0104SLCMSLCC" 'FindAccountNo(BankExportGrid.TextMatrix(Cnt, 1))
        Custcode = BankExportGrid.TextMatrix(Cnt, 1)
        ds = Replace(BankExportGrid.TextMatrix(Cnt, 6), "'", "")
        
        If BankExportGrid.TextMatrix(Cnt, 1) = "LEPL" Or BankExportGrid.TextMatrix(Cnt, 1) = "LAKSH" Then
            CompName = "LAKSHYA EDUCARE PVT LTD"
        Else
            CompName = "MT EDUCARE LIMITED"
        End If

        
    Else
        OutputStr = OutputStr & AccountNo & "    INR " & String(14 - Len(Format(BankExportGrid.TextMatrix(Cnt, 7), "0.00")), " ") & Format(BankExportGrid.TextMatrix(Cnt, 7), "0.00") & " " & Format(BankExportGrid.TextMatrix(Cnt, 1), "000000") & " " & Custcode & "/" & ds & String(23 - Len(Custcode & "/" & ds), " ") & Format(BankExportGrid.TextMatrix(Cnt, 1), "000000") & String(17 - Len(Format(BankExportGrid.TextMatrix(Cnt, 7), "0.00")), " ") & Format(BankExportGrid.TextMatrix(Cnt, 7), "0.00") & Format(Replace(BankExportGrid.TextMatrix(Cnt, 3), "'", ""), "00") & " " & BankExportGrid.TextMatrix(Cnt, 13) & "                                                                                " & CompName & vbCrLf
        'OutputStr = OutputStr & AccountNo & "    " & "INR" & vbTab & Format(BankExportGrid.TextMatrix(Cnt, 7), "000000.00") & vbTab & Format(BankExportGrid.TextMatrix(Cnt, 1), "000000") & vbTab & Custcode & "/" & ds & vbTab & Format(BankExportGrid.TextMatrix(Cnt, 1), "000000") & vbTab & Format(BankExportGrid.TextMatrix(Cnt, 7), "000000.00") & Format(Replace(BankExportGrid.TextMatrix(Cnt, 3), "'", ""), "00") & vbTab & BankExportGrid.TextMatrix(Cnt, 13) & vbCrLf
        RecCnt = RecCnt + 1
        If RecCnt = 250 Then
            'Save file and reset outputstr
            Me.txtCTS.Text = OutputStr
            txtCTS.SaveFile UploadFolder & "\" & Me.txtCMSSlipType.Text & "_" & DSNo & "_" & RunNo & ".txt", 1
            RecCnt = 0
            RunNo = RunNo + 1
            OutputStr = ""
        End If
    
    End If
     
Next
Me.txtCTS.Text = OutputStr

If Me.txtCTS.Text <> "" Then
    txtCTS.SaveFile UploadFolder & "\" & Me.txtCMSSlipType.Text & "_" & DSNo & "_" & RunNo & ".txt", 1
End If
End Sub

Private Function FindAccountNo(CompanyCode As String) As String
Dim PrintMatter As String
If CompanyCode = "MTEX" Then
    PrintMatter = "015105012577"
ElseIf CompanyCode = "MTMA" Then
    PrintMatter = "015105012578"
ElseIf CompanyCode = "MTCO" Then
    PrintMatter = "015105012579"
ElseIf CompanyCode = "MTGU" Then
    PrintMatter = "015105012580"
ElseIf CompanyCode = "MTTA" Then
    PrintMatter = "015105012581"
ElseIf CompanyCode = "MTRM" Then
    PrintMatter = "015105012582"
ElseIf CompanyCode = "MTCB" Then
    PrintMatter = "015105012583"
ElseIf CompanyCode = "MTIC" Then
    PrintMatter = "015105012584"
ElseIf CompanyCode = "MTSC" Then
    PrintMatter = "015105012585"
ElseIf CompanyCode = "MTKA" Then
    PrintMatter = "015105012587"
ElseIf CompanyCode = "MTIS" Then
    PrintMatter = "015105012588"
ElseIf CompanyCode = "LEPL" Or CompanyCode = "LAKSH" Then
    PrintMatter = "124405000213"
End If
FindAccountNo = PrintMatter

'0104SLCMSLCC

End Function


'Private Sub txtSlipNo_KeyPress(KeyAscii As Integer)
'On Error Resume Next
'If KeyAscii = 13 Then
'    Me.txtBarcodeNo.SetFocus
'    KeyAscii = 0
'    Exit Sub
'End If
'End Sub

Private Sub txtSlipNo_LostFocus()
'On Error GoTo ErrExit
'If Trim(txtSlipNo.Text) = "" Then Exit Sub
'
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
'Grid.Rows = 1
'
'rs1.Open "select cmsdate, FinalChequeCnt from ASPDC_CMS_CentreLog AC where CMSSlipNo ='" & Me.txtSlipNo.Text & "'", cn1, adOpenDynamic, adLockReadOnly
'If Not (rs1.BOF And rs1.EOF) Then
'    txtSlipDate.Text = Format(rs1.Fields("cmsdate").Value, "dd Mmm yyyy")
'    txtChequeCnt.Text = rs1.Fields("FinalChequeCnt").Value
'Else
'    MsgBox "Invalid CMS Slip Number.", vbCritical + vbOKOnly
'    txtSlipNo.SetFocus
'    Exit Sub
'End If
'rs1.Close
'
''Add all barcodes
'rs1.Open "Select CCCHQIdNo from ASPDC_DispatchSlipDetails where CMSDoneFlag =1 and CMSSlipNo ='" & Me.txtSlipNo.Text & "'", cn1, adOpenDynamic, adLockReadOnly
'If Not (rs1.BOF And rs1.EOF) Then
'    rs1.MoveFirst
'    Do While Not rs1.EOF
'        Me.lstBarcode.AddItem rs1!CCCHQIdNo
'        rs1.MoveNext
'    Loop
'End If
'rs1.Close
'
'cn1.Close
'Exit Sub
'
'ErrExit:
'MsgBox Err.Description
End Sub
