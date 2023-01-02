VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmCMSDeposit_Sub1 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CMS Deposit Slip Entry"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11940
   Icon            =   "FrmCMSDeposit_Sub1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   11940
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptOut 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Out Station Clearing"
      Height          =   195
      Left            =   5280
      TabIndex        =   23
      Top             =   540
      Width           =   1815
   End
   Begin VB.OptionButton OptLocal 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Local Clearing"
      Height          =   195
      Left            =   5280
      TabIndex        =   22
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox TxtUserName 
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
      Left            =   1800
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxtSeriesCode 
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
      Left            =   1320
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdSkip 
      BackColor       =   &H00FFFFFF&
      Caption         =   "S&kip"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin RichTextLib.RichTextBox txtCSV 
      Height          =   615
      Left            =   10920
      TabIndex        =   18
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      _Version        =   393217
      TextRTF         =   $"FrmCMSDeposit_Sub1.frx":030A
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Export to Excel"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7680
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   4560
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSlipAmt 
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
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1815
   End
   Begin VB.TextBox txtChqCnt 
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Show"
      Height          =   375
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   975
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
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Accept"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7680
      Width           =   1695
   End
   Begin VB.TextBox txtCode 
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
      Left            =   840
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
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
      Left            =   360
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "ADD"
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5370
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9472
      _Version        =   393216
      Cols            =   17
      FixedCols       =   2
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
      FormatString    =   $"FrmCMSDeposit_Sub1.frx":038C
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
   Begin MSComCtl2.DTPicker dtSlip 
      Height          =   315
      Left            =   9000
      TabIndex        =   1
      Top             =   360
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
      Format          =   126353411
      CurrentDate     =   39310
   End
   Begin VB.Label lblDepType 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Bank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   524
      Left            =   120
      TabIndex        =   24
      Top             =   960
      Width           =   1425
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Press 'F3' to Search a Particular Entry by Cheque Number"
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
      Height          =   285
      Left            =   1680
      TabIndex        =   17
      Top             =   1200
      Width           =   10095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Double Click on Cheque Entry or Press 'Spacebar' to Select or Deselect any Cheque for Deposit Slip"
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
      Height          =   285
      Left            =   1680
      TabIndex        =   16
      Top             =   960
      Width           =   10095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Total Slip Amount"
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
      Left            =   7215
      TabIndex        =   15
      Top             =   7200
      Width           =   1545
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "No. of Instruments"
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
      Left            =   585
      TabIndex        =   13
      Top             =   7200
      Width           =   1605
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   120
      Top             =   7080
      Width           =   11655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Deposit Slip Number"
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
      Left            =   270
      TabIndex        =   9
      Top             =   360
      Width           =   1755
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Deposit Slip Date"
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
      Left            =   7320
      TabIndex        =   8
      Top             =   360
      Width           =   1515
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   120
      Top             =   120
      Width           =   11655
   End
End
Attribute VB_Name = "FrmCMSDeposit_Sub1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdmnFlag As Boolean
Dim SlipId As Long
Public ModFlag As Boolean
Public CmsChkFlag As Boolean
Public PrevVal As String
Public PrevCnt As Long
Public ChqNewStatus As String

Private Sub cmdCancel_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cmdExport_Click()
'''''On Error Resume Next
'''''Dim wrkErrorMessage As String
'''''Dim wrkOutputFile As String
'''''Dim wrkProjectName As String
'''''Dim ColCnt, RowCnt As Integer
'''''
'''''wrkProjectName = "edumate - CMS Deposit Slip"
'''''
''''''Set Output File Name
'''''cmd.CancelError = True
'''''cmd.Flags = cdlOFNOverwritePrompt
'''''cmd.DefaultExt = "xls"
'''''cmd.Filter = "Excel Files|*.XLS|All files|*.*"
'''''cmd.ShowSave
'''''If Err.Number = cdlCancel Then
'''''    Exit Sub
'''''Else
'''''    wrkOutputFile = cmd.FileName
'''''End If
'''''
'''''Dim cnAim As ADODB.Connection
'''''Set cnAim = New ADODB.Connection
'''''
'''''cnAim.ConnectionString = ModInit.ConnectStringAIM
'''''cnAim.Open , , "panaaim"
'''''
'''''Dim rs1 As ADODB.Recordset
'''''Set rs1 = New ADODB.Recordset
'''''
'''''CompanyCode = ModInit.LicCode & ModInit.BranchCode
'''''rs1.Open "Select Company_Name, BankAccount_No, CenterName from CompanyDetails where CentreCode ='" & CompanyCode & "'", cnAim, adOpenDynamic, adLockReadOnly
'''''If Not (rs1.BOF And rs1.EOF) Then
'''''    CompanyName = rs1!Company_Name
'''''    BankAccount_No = rs1!BankAccount_No
'''''    CenterName = rs1!CenterName
'''''End If
'''''rs1.Close
'''''cnAim.Close
'''''
'''''ModGridToExcel.Head1 = "Deposit Slip Number"
'''''ModGridToExcel.Head2 = "Deposit Slip Date"
'''''ModGridToExcel.Head3 = "No. of Instruments"
'''''ModGridToExcel.Head4 = "Total Slip Amount"
'''''ModGridToExcel.Head5 = "Centre Name"
'''''
'''''ModGridToExcel.Head21 = Me.txtSlipNo.Text
'''''ModGridToExcel.Head22 = Format(dtSlip.Value, "ddMmyyyy")
'''''ModGridToExcel.Head23 = Me.txtChqCnt.Text
'''''ModGridToExcel.Head24 = Format(Me.txtSlipAmt.Text, "##,##,##0.00")
'''''ModGridToExcel.Head25 = CompanyName & " " & CenterName
'''''
'''''With Grid
'''''        '  Load Field Names
'''''    ModGridToExcel.NumberColumns = 60
'''''    ReDim ModGridToExcel.FieldNames(1 To 60)
'''''    For ColCnt = 1 To 10
'''''        ModGridToExcel.FieldNames(ColCnt) = .TextArray(ColCnt)
'''''    Next
'''''
'''''    '  Load Data Array
'''''    ModGridToExcel.NumberRows = .Rows - 1
'''''    ReDim ModGridToExcel.FieldData(1 To .Rows - 1, 1 To 60)
'''''    For RowCnt = 1 To .Rows - 1
'''''        For ColCnt = 1 To 10
'''''            ModGridToExcel.FieldData(RowCnt, ColCnt) = .TextMatrix(RowCnt, ColCnt)
'''''        Next
'''''    Next
'''''
'''''End With
'''''
''''''  Save the Grid as a File
'''''If (SaveExcelWorksheetCMS(wrkProjectName, wrkOutputFile, wrkErrorMessage) = False) Then
'''''    MsgBox "Error in data transfer."
'''''    Exit Sub
'''''End If


On Error Resume Next
Dim wrkErrorMessage As String
Dim wrkOutputFile As String
Dim wrkProjectName As String
Dim ColCnt, RowCnt As Integer

'Set Output File Name
cmd.CancelError = True
cmd.Flags = cdlOFNOverwritePrompt
cmd.DefaultExt = "csv"
cmd.Filter = "CSV Files|*.csv|All files|*.*"
cmd.ShowSave
If Err.Number = cdlCancel Then
    Exit Sub
Else
    wrkOutputFile = cmd.FileName
End If

Dim CnAim As ADODB.Connection
Set CnAim = New ADODB.Connection

CnAim.ConnectionString = ModInit.ConnectStringAIM
CnAim.Open , , "panaaim"

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

CompanyCode = ModInit.Liccode & ModInit.BranchCode
rs1.Open "Select Company_Name, BankAccount_No, CenterName from CompanyDetails where CentreCode ='" & CompanyCode & "'", CnAim, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    CompanyName = rs1!Company_Name
    BankAccount_No = rs1!BankAccount_No
    CenterName = rs1!CenterName
End If
rs1.Close
CnAim.Close

ModGridToExcel.Head1 = "Deposit Slip Number"
ModGridToExcel.Head2 = "Deposit Slip Date"
ModGridToExcel.Head3 = "No. of Instruments"
ModGridToExcel.Head4 = "Total Slip Amount"
ModGridToExcel.Head5 = "Centre Name"

ModGridToExcel.Head21 = Me.txtSlipNo.Text
ModGridToExcel.Head22 = Format(dtSlip.Value, "ddMmyyyy")
ModGridToExcel.Head23 = Me.txtChqCnt.Text
ModGridToExcel.Head24 = Format(Me.txtSlipAmt.Text, "##,##,##0.00")
ModGridToExcel.Head25 = CompanyName & " " & CenterName

With Grid
        '  Load Field Names
    ModGridToExcel.NumberColumns = 10
    ReDim ModGridToExcel.FieldNames(1 To 60)
    For ColCnt = 1 To 10
        ModGridToExcel.FieldNames(ColCnt) = .TextArray(ColCnt)
    Next

    '  Load Data Array
    ModGridToExcel.NumberRows = .Rows - 1
    ReDim ModGridToExcel.FieldData(1 To .Rows - 1, 1 To 60)
    For RowCnt = 1 To .Rows - 1
        For ColCnt = 1 To 10
            ModGridToExcel.FieldData(RowCnt, ColCnt) = .TextMatrix(RowCnt, ColCnt)
        Next
    Next

End With

'call function to convert
txtCSV.Text = ModGridToExcel.CSV_Output_CMS()

txtCSV.SaveFile wrkOutputFile, 1
MsgBox "CSV file created at " & wrkOutputFile, vbInformation + vbOKOnly

End Sub

Private Sub cmdPrint_Click()
On Error Resume Next
Dim res As Integer
Dim CnAim As ADODB.Connection
Set CnAim = New ADODB.Connection

CnAim.ConnectionString = ModInit.ConnectStringAIM
CnAim.Open , , "panaaim"

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset


If cmdSave.Enabled = True Then
    res = MsgBox("This action will finalise the Deposit Slip entry and you will not be able to Edit any entry in it.  Do you want to proceed now?", vbQuestion + vbYesNo)
    If res = 7 Then Exit Sub

    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    Dim rs3 As ADODB.Recordset
    Set rs3 = New ADODB.Recordset

    Dim cnYr As New ADODB.Connection

    'Make  Flag as Cleared in each cheque
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject

    Dim YrCStr As String

    rs1.Open "Select * from DepositSlipDetails where SlipCode =" & txtCode.Text & " order by entrycode", CnAim, adOpenDynamic, adLockReadOnly
    If Not (rs1.BOF And rs1.EOF) Then
        rs1.MoveFirst
        Do While Not rs1.EOF
            'Open yearfile
            YrName = rs1!YearName


            If YrName <> "" Then
                'Open the year file
                YrPath = App.Path & "\smf" & StrReverse(Left(YrName, 4)) & "yr1.scd"

                If fso.FileExists(YrPath) Then

                    YrCStr = "Provider=MSDASQL.1;Persist Security Info=False;User ID=admin;Extended Properties=""DBQ=" & YrPath & ";DefaultDir= " & App.Path & ";Driver={Microsoft Access Driver (*.mdb)};DriverId=281;FIL=MS Access;FILEDSN=" & YrPath & ";MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"";Initial Catalog=" & YrPath & """"

                    cnYr.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;User ID=admin;Extended Properties=""DBQ=" & YrPath & ";DefaultDir= " & App.Path & ";Driver={Microsoft Access Driver (*.mdb)};DriverId=281;FIL=MS Access;FILEDSN=" & YrPath & ";MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"";Initial Catalog=" & YrPath & """"
                    cnYr.Open , , "panatech"

                    rs2.Open "Alter table StudentPayment Add Column SlipNo Text(50) Null", cnYr, adOpenDynamic, adLockPessimistic
                    rs2.Open "Alter table StudentOtherPayment Add Column SlipNo Text(50) Null", cnYr, adOpenDynamic, adLockPessimistic


                    If Left(rs1!RcptCode, 1) = "O" Then
                        'Other Payment Entries
                        rs2.Open "Select * from StudentOtherPayment where ChkNo ='" & rs1!ChequeNo & "' and RcptCode ='" & Replace(rs1!RcptCode, "O", "") & "'", cnYr, adOpenDynamic, adLockPessimistic
                        If Not (rs2.BOF And rs2.EOF) Then
    '                        rs2.Fields("ChkStatus").Value = "Deposited"

                            If Me.lblDepType.Caption = "CO" Then
                                rs2.Fields("ChkStatus").Value = "In transit"
                            Else
                                rs2.Fields("ChkStatus").Value = "Deposited"
                            End If
                            rs2.Fields("DepositDate").Value = Me.dtSlip.Value
                            rs2.Fields("RealDate").Value = Me.dtSlip.Value
                            rs2.Fields("UpLoadFlag").Value = "False"
                            rs2.Fields("SlipNo").Value = Trim(txtSlipNo.Text)
                            rs2.Update
                        End If
                        rs2.Close
                    Else
                        rs2.Open "Select * from StudentPayment where RecordDelFlag = 0 and ChkNo ='" & rs1!ChequeNo & "' and RcptCode ='" & rs1!RcptCode & "'", cnYr, adOpenDynamic, adLockPessimistic
                        If Not (rs2.BOF And rs2.EOF) Then
                            If Me.lblDepType.Caption = "CO" Then
                                rs2.Fields("ChkStatus").Value = "In transit"
                            Else
                                rs2.Fields("ChkStatus").Value = "Deposited"
                            End If
                            rs2.Fields("DepositDate").Value = Me.dtSlip.Value
                            rs2.Fields("RealDate").Value = Me.dtSlip.Value
                            rs2.Fields("SlipNo").Value = Trim(txtSlipNo.Text)
                            rs2.Fields("UpLoadFlag").Value = "False"
                            rs2.Update
                        End If
                        rs2.Close
                    End If
                    cnYr.Close

                    ModInit.UpdateStudentLedger rs1!SBEntryCode, YrCStr

                End If
            Else
                'Cash deposit to DD entry
                If Left(rs1!RcptCode, 2) = "CS" Then
                    'Do nothing

                End If

            End If

            rs1.MoveNext
        Loop
    End If
    rs1.Close

    rs1.Open "Select * from DepositSlip where SlipCode = " & txtCode.Text, CnAim, adOpenDynamic, adLockOptimistic
    If Not (rs1.BOF And rs1.EOF) Then
        rs1.Fields("PrintFlag") = 1
        rs1.Update
    End If
    rs1.Close

    CompanyCode = ModInit.Liccode & FrmCMSDeposit.cboInstCode.Text
    rs1.Open "Select Company_Name, BankAccount_No, CenterName from CompanyDetails where CentreCode ='" & CompanyCode & "'", CnAim, adOpenDynamic, adLockReadOnly
    If Not (rs1.BOF And rs1.EOF) Then
        CompanyName = rs1!Company_Name
        CenterName = rs1!CenterName
    End If
    rs1.Close

    AddEventE14 txtSlipNo.Text

    ModComFunction.FillDayLogBook "CMS Deposit Slip", " Deposit Slip No - " & Me.txtSlipNo.Text & " (Date -" & Format(dtSlip.Value, "dd Mmm yyyy") & ")"

    cmdSave.Enabled = False
    cmdExport.Enabled = True

    FrmMenu_sub_Logo.FindChequesForClearing
    CmdShow_Click

'    Unload Me


Else
    CompanyCode = ModInit.Liccode & ModInit.BranchCode
    rs1.Open "Select Company_Name, BankAccount_No, CenterName from CompanyDetails where CentreCode ='" & CompanyCode & "'", CnAim, adOpenDynamic, adLockReadOnly
    If Not (rs1.BOF And rs1.EOF) Then
        CompanyName = rs1!Company_Name
        BankAccount_No = rs1!BankAccount_No
        CenterName = rs1!CenterName
    End If
    rs1.Close
End If
CnAim.Close

For Cnt = Grid.Rows - 1 To 1 Step -1
    If Grid.TextMatrix(Cnt, 0) = "" Then
        If Grid.Rows > 2 Then
            Grid.RemoveItem Cnt
        Else
            Grid.Rows = 1
        End If
    End If
Next

PrintRes
Exit Sub

'Start Print Code
'Print from student grid
Dim TROWS, DPAGES, IPAGES, TPAGES, ENDNUM As Long
Dim intCopies, NumCopies As Integer

Dim X As Integer
Dim Y As Integer
Dim nxtrow As Integer
Dim LeftMargin As Integer
Dim RightMargin As Long
Dim TopMargin As Long
Dim BottomMargin As Long
Dim YPos, YStart As Long

'For A4 size paper
LeftMargin = 500
RightMargin = 16000
TopMargin = 400
BottomMargin = 9000

' Set Cancel to True.
cmd.CancelError = True

On Error GoTo ErrHandler

'Set flags
cmd.Flags = &H100000 Or &H4

' Display the Print dialog box.
cmd.ShowPrinter

' Get user-selected values from the dialog box
NumCopies = cmd.Copies
Printer.Orientation = cdlLandscape
PrintLines = 25

'Determines Number of Records divisible by 30 and remainder

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

'Setup Global for printer
Printer.Font = "arial"
Printer.DrawWidth = 2
Printer.FillStyle = vbFSSolid

'Start copies loop
For intCopies = 1 To NumCopies
    'Begin loop for pages

    For X = 1 To TPAGES
        YPos = TopMargin


        Printer.FontBold = False
        Printer.FontSize = 10
        Printer.FontName = "Arial"

        Printer.Line (LeftMargin, YPos)-(RightMargin, YPos + 15)
        YStart = YPos

        YPos = YPos + 100

        Dim PrintMatter As String

        PrintMatter = "Deposit Slip Number"
        Printer.CurrentX = LeftMargin + (2 * 567) - (Printer.TextWidth(PrintMatter) / 2)
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        PrintMatter = "Deposit Slip Date"
        Printer.CurrentX = LeftMargin + (6 * 567) - (Printer.TextWidth(PrintMatter) / 2)
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        PrintMatter = "No. of Instruments"
        Printer.CurrentX = LeftMargin + (10 * 567) - (Printer.TextWidth(PrintMatter) / 2)
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        PrintMatter = "Total Slip Amount"
        Printer.CurrentX = LeftMargin + (14 * 567) - (Printer.TextWidth(PrintMatter) / 2)
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        PrintMatter = "Centre Name"
        Printer.CurrentX = LeftMargin + (16 * 567) + (RightMargin - (LeftMargin + 16 * 567)) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        YPos = YPos + Printer.TextHeight("X") + 100
        Printer.Line (LeftMargin, YPos)-(RightMargin, YPos + 5)
        YPos = YPos + 100

        PrintMatter = Me.txtSlipNo.Text
        Printer.CurrentX = LeftMargin + (2 * 567) - (Printer.TextWidth(PrintMatter) / 2)
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        PrintMatter = Format(dtSlip.Value, "ddMmyyyy")
        Printer.CurrentX = LeftMargin + (6 * 567) - (Printer.TextWidth(PrintMatter) / 2)
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        PrintMatter = Me.txtChqCnt.Text
        Printer.CurrentX = LeftMargin + (10 * 567) - (Printer.TextWidth(PrintMatter) / 2)
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        PrintMatter = Format(Me.txtSlipAmt.Text, "##,##,##0.00")
        Printer.CurrentX = LeftMargin + (14 * 567) - (Printer.TextWidth(PrintMatter) / 2)
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        PrintMatter = CompanyName & " " & CenterName
        Printer.CurrentX = LeftMargin + (16 * 567) + (RightMargin - (LeftMargin + 16 * 567)) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        YPos = YPos + Printer.TextHeight("X") + 50
        Printer.Line (LeftMargin, YPos)-(RightMargin, YPos + 5)
        Printer.Line (LeftMargin, YStart)-(LeftMargin + 2, YPos)
        Printer.Line (LeftMargin + (4 * 567), YStart)-(LeftMargin + (4 * 567) + 2, YPos)
        Printer.Line (LeftMargin + (8 * 567), YStart)-(LeftMargin + (8 * 567) + 2, YPos)
        Printer.Line (LeftMargin + (12 * 567), YStart)-(LeftMargin + (12 * 567) + 2, YPos)
        Printer.Line (LeftMargin + (16 * 567), YStart)-(LeftMargin + (16 * 567) + 2, YPos)
        Printer.Line (RightMargin, YStart)-(RightMargin + 5, YPos)

        YPos = YPos + 200
        YStart = YPos
        Printer.Line (LeftMargin, YPos)-(RightMargin, YPos + 5)

        Dim Col1Left, Col2Left, Col3Left, Col4Left, Col5Left, Col6Left, Col7Left, Col8Left, Col9Left, Col10Left As Long
        'Print column headings
        Col1Left = LeftMargin + 1300
        Col2Left = Col1Left + 1300
        Col3Left = Col2Left + 1300
        Col4Left = Col3Left + 1100
        Col5Left = Col4Left + 1100
        Col6Left = Col5Left + 1300
        Col7Left = Col6Left + 3500
        Col8Left = Col7Left + 2000
        Col9Left = Col8Left + 1000
        Col10Left = Col9Left + 1000

        YPos = YPos + 100

        Printer.CurrentY = YPos
        PrintMatter = "Instrument"
        Printer.CurrentX = LeftMargin + (Col1Left - LeftMargin) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos + Printer.TextHeight("X") + 30
        PrintMatter = "Number"
        Printer.CurrentX = LeftMargin + (Col1Left - LeftMargin) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos
        PrintMatter = "Instrument"
        Printer.CurrentX = Col1Left + (Col2Left - Col1Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos + Printer.TextHeight("X") + 30
        PrintMatter = "Date"
        Printer.CurrentX = Col1Left + (Col2Left - Col1Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos
        PrintMatter = "Instrument"
        Printer.CurrentX = Col2Left + (Col3Left - Col2Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos + Printer.TextHeight("X") + 30
        PrintMatter = "Amount"
        Printer.CurrentX = Col2Left + (Col3Left - Col2Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos
        PrintMatter = "Drawn On"
        Printer.CurrentX = Col3Left + (Col4Left - Col3Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos + Printer.TextHeight("X") + 30
        PrintMatter = "Bank"
        Printer.CurrentX = Col3Left + (Col4Left - Col3Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos
        PrintMatter = "Drawn On"
        Printer.CurrentX = Col4Left + (Col5Left - Col4Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos + Printer.TextHeight("X") + 30
        PrintMatter = "Branch"
        Printer.CurrentX = Col4Left + (Col5Left - Col4Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos
        PrintMatter = "Instrument"
        Printer.CurrentX = Col5Left + (Col6Left - Col5Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos + Printer.TextHeight("X") + 30
        PrintMatter = "Type"
        Printer.CurrentX = Col5Left + (Col6Left - Col5Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos
        PrintMatter = "Name of the Student"
        Printer.CurrentX = Col6Left + 100 ' (Col7Left - Col6Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos
        PrintMatter = "Course"
        Printer.CurrentX = Col7Left + 100 '(Col8Left - Col7Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos
        PrintMatter = "Form No"
        Printer.CurrentX = Col8Left + 100 '(Col9Left - Col8Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos
        PrintMatter = "Academic Year"
        Printer.CurrentX = Col9Left + 100 '(RightMargin - Col9Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        YPos = YPos + 2 * Printer.TextWidth("X") + 250

        'Horz divider (bottom of column titles)
        Printer.Line (LeftMargin, YPos)-(RightMargin, YPos + 5)
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
                nxtrow = 350 * (Y - (PrintLines * (X - 1)))
            Else
                nxtrow = 350 * Y
            End If

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 1)
            Printer.CurrentX = LeftMargin + (Col1Left - LeftMargin) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 2)
            Printer.CurrentX = Col1Left + (Col2Left - Col1Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 3)
            Printer.CurrentX = Col2Left + (Col3Left - Col2Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 4)
            Printer.CurrentX = Col3Left + (Col4Left - Col3Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 5)
            Printer.CurrentX = Col4Left + (Col5Left - Col4Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 6)
            Printer.CurrentX = Col5Left + (Col6Left - Col5Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 7)
            Printer.CurrentX = Col6Left + 100 '+ (Col7Left - Col6Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 8)
            Printer.CurrentX = Col7Left + 100 '+ (Col8Left - Col7Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 9)
            Printer.CurrentX = Col8Left + 100 '+ (Col9Left - Col8Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 10)
            Printer.CurrentX = Col9Left + 100 '+ (Col9Left - Col8Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter


        Next Y

        YPos = YPos + nxtrow + Printer.TextHeight("X") + 100

        Printer.Line (LeftMargin, YPos)-(RightMargin, YPos + 5)

        Printer.Line (LeftMargin, YStart)-(LeftMargin + 5, YPos)
        Printer.Line (Col1Left, YStart)-(Col1Left + 5, YPos)
        Printer.Line (Col2Left, YStart)-(Col2Left + 5, YPos)
        Printer.Line (Col3Left, YStart)-(Col3Left + 5, YPos)
        Printer.Line (Col4Left, YStart)-(Col4Left + 5, YPos)
        Printer.Line (Col5Left, YStart)-(Col5Left + 5, YPos)
        Printer.Line (Col6Left, YStart)-(Col6Left + 5, YPos)
        Printer.Line (Col7Left, YStart)-(Col7Left + 5, YPos)
        Printer.Line (Col8Left, YStart)-(Col8Left + 5, YPos)
        Printer.Line (Col9Left, YStart)-(Col9Left + 5, YPos)

        Printer.Line (RightMargin, YStart)-(RightMargin - 5, YPos)

        YPos = YPos + 100
        Printer.CurrentY = YPos
        Printer.CurrentX = LeftMargin
        Printer.Print "edumate " & Format(Date, "dd Mmm yyyy") & " - Sheet " & Printer.Page; " of " & TPAGES

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


Private Sub PrintRes()
On Error GoTo ErrHandler
Dim TROWS, DPAGES, IPAGES, TPAGES, ENDNUM As Long
Dim intCopies, NumCopies As Integer

Dim X As Integer
Dim Y As Integer
Dim nxtrow As Integer
Dim LeftMargin As Integer
Dim RightMargin As Long
Dim TopMargin As Long
Dim BottomMargin As Long
Dim YPos, YStart As Long
Dim NoOfCopies As Long

Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim cn1 As New ADODB.Connection

Dim CnAim As New ADODB.Connection
CnAim.Open ModInit.ConnectStringAIM, , "panaaim"

Dim RecordCnt As Integer
RecordCnt = 0
rs1.Open "Select count(*) as RecordCnt from DepositSlipDetails where SlipCode =" & txtCode.Text & "", CnAim, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    RecordCnt = rs1!RecordCnt
End If
rs1.Close
CnAim.Close

If RecordCnt = 0 Then Exit Sub

'For A4 size paper
LeftMargin = 200
RightMargin = 15700
TopMargin = 400
BottomMargin = 11000

' Set Cancel to True.
cmd.CancelError = True

On Error GoTo ErrHandler

'Set flags
cmd.Flags = &H100000 Or &H4

' Display the Print dialog box.
cmd.ShowPrinter

' Get user-selected values from the dialog box
NumCopies = cmd.Copies
Printer.Orientation = cdlLandscape
PrintLines = 20

'Determines Number of Records divisible by 30 and remainder

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

If OptLocal.Value = True Then
    NoOfCopies = 4
Else
    NoOfCopies = 3
End If

'Setup Global for printer
Printer.Font = "arial"
Printer.DrawWidth = 2
Printer.FillStyle = vbFSSolid

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
        If lblDepType.Caption = "Bank" Then
            If OptLocal.Value = True Then
                PrintMatter = "CITICLEAR DEPOSIT SLIP/DETAILS"
            Else
                PrintMatter = "CITICHECK DEPOSIT SLIP/DETAILS"
            End If
        Else
            PrintMatter = "CENTRE TO CO CHEQUE DEPOSIT SLIP/DETAILS"
        End If
        Printer.CurrentX = LeftMargin + 100
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        If lblDepType.Caption = "Bank" Then
            PrintMatter = "CITIBANK"
        Else
            PrintMatter = ""
        End If

        Printer.CurrentX = RightMargin - Printer.TextWidth(PrintMatter) - 100
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        YPos = YPos + Printer.TextHeight("X") + 100
        Printer.Line (LeftMargin, YPos)-(RightMargin, YPos)
        YMain = YPos

        YPos = YPos + 100
        Col1Left = LeftMargin + 3000
        Col2Left = Col1Left + 5000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        PrintMatter = "SLIP No"
        Printer.CurrentX = LeftMargin + (Col1Left - LeftMargin) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        ''''Print Slip No
        PrintMatter = txtSlipNo.Text
        Printer.CurrentX = LeftMargin + (Col1Left - LeftMargin) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.CurrentY = YPos + Printer.TextHeight("X") + 550
        Printer.FontSize = 12
        Printer.Print PrintMatter
        Printer.FontSize = 10

        PrintMatter = "(Mandatory Field)"
        Printer.CurrentX = Col1Left + (Col2Left - Col1Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        Printer.FontBold = False
        PrintMatter = "Customer's Name & Divn."
        Printer.CurrentX = Col1Left + (Col2Left - Col1Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.CurrentY = YPos + Printer.TextHeight("X") + 50
        Printer.Print PrintMatter

        PrintMatter = "Date"
        Printer.CurrentX = Col2Left + 250
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        ''''''''''''''''Printing the Date Value and Gross Amount
        PrintMatter = Format(Now, "dd MMM yyyy")
        Printer.CurrentX = Col2Left + 250
        Printer.CurrentY = YPos + Printer.TextHeight("X") + 200
        Printer.Print PrintMatter

        PrintMatter = txtSlipAmt.Text
        Printer.CurrentX = RightMargin - Printer.TextWidth(PrintMatter) - 1200
        Printer.CurrentY = YPos + Printer.TextHeight("X") + 200
        Printer.Print PrintMatter
        '''''''''''''''''''''''''''''''''''''

        PrintMatter = "Gross Deposit Amount"
        Printer.CurrentX = RightMargin - Printer.TextWidth(PrintMatter) - 1200
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        Printer.Line (Col1Left, YPos + Printer.TextHeight("X") + 550)-(RightMargin, YPos + Printer.TextHeight("X") + 550)
        YDum = YPos + Printer.TextHeight("X") + 550
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim DiffVal As Single
        DiffVal = Val(Val(Val(RightMargin) - Val(Col2Left)) / 3)

        YPos = YPos + Printer.TextHeight("X") + 600
        PrintMatter = "Pickup Location"
        Printer.CurrentX = Col2Left + 100
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        Printer.CurrentX = Col2Left + Val(Val(DiffVal) * 1) + 100
        PXval = Printer.CurrentX
        PrintMatter = "No of Checks"
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        '''''''''''''Printing No of No of Checks''''''''''''
        Printer.CurrentX = Col2Left + Val(Val(DiffVal) * 1) + 250

        PXval = Printer.CurrentX
        PrintMatter = Val(txtChqCnt.Text)
        Printer.CurrentY = YPos + Printer.TextHeight("X") + 200
        Printer.Print PrintMatter

        Printer.CurrentX = Col1Left + 100 'Printing Customer Code''''''''''''
        PrintMatter = FrmCMSDeposit.cboCCode.Text   ' ModInit.CCode
        Printer.CurrentY = YPos + Printer.TextHeight("X") + 200
        Printer.Print PrintMatter

        Printer.CurrentX = Col1Left + Val(Val(Col2Left - Col1Left) / 2) + 100 'Printing PickUpPoint''''''''''''
        PrintMatter = FrmCMSDeposit.cboPPoint.Text   'ModInit.PPoint
        Printer.CurrentY = YPos + Printer.TextHeight("X") + 200
        Printer.Print PrintMatter

        Printer.CurrentX = Col2Left + 100 'Printing Pickup Location''''''''''''
        PrintMatter = FrmCMSDeposit.cboPLoc.Text   ' ModInit.PLoc
        Printer.CurrentY = YPos + Printer.TextHeight("X") + 200
        Printer.Print PrintMatter

        Printer.CurrentX = Col2Left + Val(Val(DiffVal) * 2) + 100 'Printing Customer's Ref''''''''''''
        PrintMatter = FrmCMSDeposit.cboCref.Text '  ModInit.Cref
        Printer.CurrentY = YPos + Printer.TextHeight("X") + 200
        Printer.Print PrintMatter

        '''''''''''''''''''''''''''''''''''''''''''
        Printer.CurrentX = Col2Left + Val(Val(DiffVal) * 2) + 100
        PrintMatter = "Customer's Ref"
        Printer.CurrentY = YPos
        Printer.Print PrintMatter
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        YPos = YPos + Printer.TextHeight("X") + 600

        PrintMatter = "Customer's Code"
        Printer.CurrentX = Col1Left + 100
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        Printer.CurrentX = Col1Left + Val(Val(Col2Left - Col1Left) / 2) + 100 'Col2Left - Printer.TextWidth(PrintMatter) - 250
        PrintMatter = "PickUpPoint"
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        YPos = YPos + Printer.TextHeight("X") + 50
        YStart1 = YPos

        Printer.Line (Col2Left + Val(Val(DiffVal) * 1), YDum)-(Col2Left + Val(Val(DiffVal) * 1), YPos - 100)
        Printer.Line (Col2Left + Val(Val(DiffVal) * 2), YDum)-(Col2Left + Val(Val(DiffVal) * 2), YPos - 100)
        Printer.Line (Col2Left, YPos - 100)-(RightMargin, YPos - 100)

        Printer.Line (LeftMargin, YMain)-(LeftMargin, YPos)
        Printer.Line (Col1Left, YMain)-(Col1Left, YPos)
        Printer.Line (Col2Left, YMain)-(Col2Left, YPos)
        Printer.Line (RightMargin, YMain)-(RightMargin, YPos)


        Printer.Line (Col1Left + Val(Val(Col2Left - Col1Left) / 2), YDum)-(Col1Left + Val(Val(Col2Left - Col1Left) / 2), YPos)

        Printer.Line (LeftMargin, YPos)-(RightMargin, YPos)
        Printer.FontSize = 8
        'Print column headings
        Col1Left = LeftMargin + 500
        Col2Left = Col1Left + 2000
        Col3Left = Col2Left + 2500
        Col4Left = Col3Left + 3500
        Col5Left = Col4Left + 5000
        Col6Left = RightMargin - 500

        YPos = YPos + 50

        Printer.CurrentY = YPos
        PrintMatter = "No."
        Printer.CurrentX = LeftMargin + (Col1Left - LeftMargin) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos
        PrintMatter = "Check No."
        Printer.CurrentX = Col1Left + (Col2Left - Col1Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos
        PrintMatter = "Check Date"
        Printer.CurrentX = Col2Left + (Col3Left - Col2Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos
        PrintMatter = "Drawer"
        Printer.CurrentX = Col3Left + (Col4Left - Col3Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos
        PrintMatter = "Drawee Bank"
        Printer.CurrentX = Col4Left + (Col5Left - Col4Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

        Printer.CurrentY = YPos
        PrintMatter = "Amount"
        Printer.CurrentX = Col5Left + (RightMargin - Col5Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
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
                nxtrow = 350 * (Y - (PrintLines * (X - 1)))
            Else
                nxtrow = 350 * Y
            End If

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Val(Y)
            Printer.CurrentX = LeftMargin + (Col1Left - LeftMargin) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 1)
            Printer.CurrentX = Col1Left + (Col2Left - Col1Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 2)
            Printer.CurrentX = Col2Left + (Col3Left - Col2Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 7)
            Printer.CurrentX = Col3Left + 100
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Left(Grid.TextMatrix(Y, 15), 45) 'Left(Grid.TextMatrix(Y, 15), 40)
            Printer.CurrentX = Col4Left + 100
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Grid.TextMatrix(Y, 3)
            Printer.CurrentX = RightMargin - (Printer.TextWidth(PrintMatter)) - 150
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
'        Printer.Line (Col6Left, YStart1)-(Col6Left, YPos)
        Printer.Line (LeftMargin, YStart)-(LeftMargin, YPos)
        Printer.Line (RightMargin, YStart)-(RightMargin, YPos)

        YPos = YPos + 100

        If X = TPAGES Then

            Printer.CurrentY = YPos
            If lblDepType.Caption = "Bank" Then
                PrintMatter = "Please ensure all checks are drawn in favour of 'CITIBANK N.A. A/C COMPANY NAME' "
            Else
                PrintMatter = "Please ensure that checks are submitted along with this slip."
            End If
            Printer.CurrentX = LeftMargin + 100
            Printer.Print PrintMatter

            Printer.CurrentY = YPos
            PrintMatter = "TOTAL ->"
            Printer.CurrentX = Col5Left - Printer.TextWidth(PrintMatter) - 150
            Printer.Print PrintMatter

            Printer.CurrentY = YPos
            PrintMatter = Format(txtSlipAmt.Text, "#,##,##,##0.00")
            Printer.CurrentX = RightMargin - Printer.TextWidth(PrintMatter) - 150
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

        If OptLocal.Value = True Then
            If intCopies = 4 Then
                PrintMatter = "Customer Copy"
            ElseIf intCopies = 3 Then
                PrintMatter = "Co-Ordinator Copy"
            Else
                PrintMatter = "Citibank Copy"
            End If
        Else
            If intCopies = 3 Then
                PrintMatter = "Customer Copy"
            ElseIf intCopies = 2 Then
                PrintMatter = "Co-Ordinator Copy"
            Else
                PrintMatter = "Citibank Copy"
            End If
        End If

        PrintMatter1 = "(Page of " & Val(X) & " of " & Val(TPAGES) & " )"

        Printer.CurrentY = YPos
        Printer.CurrentX = LeftMargin + ((RightMargin - LeftMargin) / 2) - (Printer.TextWidth(PrintMatter) / 2)
        Printer.Print PrintMatter

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

Private Sub cmdSave_Click()
On Error Resume Next
TxtUserName.Text = ""
If checkvalid = True Then
    Dim cn1 As ADODB.Connection
    Set cn1 = New ADODB.Connection
    cn1.ConnectionString = ModInit.ConnectStringAIM
    cn1.Open , , "panaaim"

    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    Dim str As String

    FrmCMSDeposit_Sub3.Show vbModal

    If Trim(TxtUserName.Text) = "" Then
        MsgBox "Enter UserName.", vbCritical + vbOKOnly
        Exit Sub
    End If
    If txtFlag.Text = "ADD" Then
        ReadNewCode
        str = "Select * from DepositSlip where SlipNo ='" & Trim(UCase(Me.txtSlipNo.Text)) & "'"
        rs1.Open str, cn1, adOpenDynamic, adLockOptimistic
        If rs1.BOF And rs1.EOF Then
            AddSlip
        Else
            MsgBox "The Deposit Slip Number already exists.", vbInformation + vbOKOnly, "Error"
            txtSlipNo.SetFocus
            rs1.Close
            Exit Sub
        End If
        rs1.Close
        cn1.Close

        txtCode.Text = SlipId
    Else
        str = "Select * from DepositSlip where SlipNo ='" & Trim(UCase(Me.txtSlipNo.Text)) & "' and SlipCode <> " & txtCode.Text
        rs1.Open str, cn1, adOpenDynamic, adLockPessimistic
        If rs1.BOF And rs1.EOF Then
            ModifySlip
        Else
            MsgBox "The Deposit Slip Number already exists.", vbInformation + vbOKOnly, "Error"
            txtSlipNo.SetFocus
            rs1.Close
            Exit Sub
        End If
        rs1.Close
        cn1.Close
    End If
    ModFlag = True

    'cmdPrint.Caption = "Confirm"
    cmdPrint.Enabled = True
    cmdExport.Enabled = False
    cmdPrint.SetFocus
End If
End Sub

Private Sub AddSlip()
On Error Resume Next
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection
cn1.ConnectionString = ModInit.ConnectStringAIM
cn1.Open , , "panaaim"

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

rs1.Open "Alter table DepositSlipDetails add Column SBEntryCode Text(50) Null", cn1, adOpenDynamic, adLockPessimistic


rs1.Open "Select * from DepositSlip", cn1, adOpenDynamic, adLockOptimistic
rs1.AddNew
rs1.Fields("SlipCode").Value = SlipId
rs1.Fields("SlipNo").Value = Trim(txtSlipNo.Text)
rs1.Fields("SlipDate").Value = dtSlip.Value
rs1.Fields("ChqCnt").Value = Val(Me.txtChqCnt.Text)
rs1.Fields("ChqAmt").Value = Val(Replace(Me.txtSlipAmt.Text, ",", ""))
rs1.Fields("PrintFlag").Value = 0
rs1.Fields("UploadFlag").Value = "False"
rs1.Fields("SeriesCode").Value = Val(TxtSeriesCode.Text)
rs1.Fields("UserName").Value = TxtUserName.Text
rs1.Fields("DepositSlipType").Value = Me.lblDepType.Caption
rs1.Fields("MISInstituteCode").Value = FrmCMSDeposit.cboInstCode.Text
If Me.OptLocal.Value = True Then
    rs1!ClearingType = 0
Else
    rs1!ClearingType = 1
End If
rs1.Update
rs1.Close

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

'save deposit slip details
rs1.Open "Delete from DepositSlipDetails where SlipCode =" & SlipId, cn1, adOpenDynamic, adLockPessimistic

Dim CCnt As Integer
CCnt = 0
rs1.Open "Select * from DepositSlipDetails where SlipCode =" & SlipId, cn1, adOpenDynamic, adLockPessimistic
For Cnt = 1 To Grid.Rows - 1
    If Grid.TextMatrix(Cnt, 0) = ChqNewStatus Then
        CCnt = CCnt + 1
        rs1.AddNew
        rs1.Fields("SlipCode").Value = SlipId
        rs1.Fields("EntryCode").Value = CCnt
        rs1.Fields("YearName").Value = Grid.TextMatrix(Cnt, 10)
        rs1.Fields("BatchCode").Value = Grid.TextMatrix(Cnt, 11)
        rs1.Fields("RollNo").Value = Grid.TextMatrix(Cnt, 12)
        rs1.Fields("ChequeNo").Value = Grid.TextMatrix(Cnt, 1)

        rs1.Fields("ChequeDate").Value = Format(Grid.TextMatrix(Cnt, 14), "dd Mmm yyyy") 'Format(Left(Grid.TextMatrix(Cnt, 2), 2) & "-" & Mid(Grid.TextMatrix(Cnt, 2), 3, 2) & "-" & Right(Grid.TextMatrix(Cnt, 2), 4), "dd Mmm yyyy")

        rs1.Fields("ChequeAmt").Value = Grid.TextMatrix(Cnt, 3)
        rs1.Fields("RcptCode").Value = Grid.TextMatrix(Cnt, 13)
        rs1.Fields("SBEntryCode").Value = Grid.TextMatrix(Cnt, 16)

        'For DD Generated from Cash Deposit Slip
        If Left(Grid.TextMatrix(Cnt, 13), 2) = "CS" Then
            'Find that Cash Deposit Slip and change

            rs2.Open "Update DepositSlip set DDToCMSFlag = 1, DDToCMSSlipCode = " & SlipId & " where SlipCode = " & Right(Grid.TextMatrix(Cnt, 13), Len(Grid.TextMatrix(Cnt, 13)) - 2) & " and CashFlag =1 and CashToOptFlag = 1", cn1, adOpenDynamic, adLockPessimistic
        End If

        rs1.Update
    End If
Next
rs1.Close
cn1.Close
End Sub

Private Sub ModifySlip()
On Error Resume Next
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection
cn1.ConnectionString = ModInit.ConnectStringAIM
cn1.Open , , "panaaim"

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

rs1.Open "Select * from DepositSlip where SlipCode = " & txtCode.Text, cn1, adOpenDynamic, adLockOptimistic
If rs1.BOF And rs1.EOF Then rs1.AddNew

rs1.Fields("SlipCode") = txtCode.Text
rs1.Fields("SlipNo").Value = Trim(txtSlipNo.Text)
rs1.Fields("SlipDate").Value = dtSlip.Value
rs1.Fields("ChqCnt").Value = Val(Me.txtChqCnt.Text)
rs1.Fields("ChqAmt").Value = Val(Replace(Me.txtSlipAmt.Text, ",", ""))
rs1.Fields("UploadFlag").Value = "False"
rs1.Fields("SeriesCode").Value = Val(TxtSeriesCode.Text)
rs1.Fields("UserName").Value = TxtUserName.Text
If Me.OptLocal.Value = True Then
    rs1!ClearingType = 0
Else
    rs1!ClearingType = 1
End If

rs1.Update
rs1.Close

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

'save deposit slip details
rs1.Open "Delete from DepositSlipDetails where SlipCode =" & txtCode.Text, cn1, adOpenDynamic, adLockPessimistic

Dim CCnt As Integer
CCnt = 0
rs1.Open "Select * from DepositSlipDetails where SlipCode =" & txtCode.Text, cn1, adOpenDynamic, adLockPessimistic
For Cnt = 1 To Grid.Rows - 1
    If Grid.TextMatrix(Cnt, 0) = ChqNewStatus Then
        CCnt = CCnt + 1
        rs1.AddNew
        rs1.Fields("SlipCode").Value = txtCode.Text
        rs1.Fields("EntryCode").Value = CCnt
        rs1.Fields("YearName").Value = Grid.TextMatrix(Cnt, 10)
        rs1.Fields("BatchCode").Value = Grid.TextMatrix(Cnt, 11)
        rs1.Fields("RollNo").Value = Grid.TextMatrix(Cnt, 12)
        rs1.Fields("ChequeNo").Value = Grid.TextMatrix(Cnt, 1)

        rs1.Fields("ChequeDate").Value = Format(Grid.TextMatrix(Cnt, 14), "dd Mmm yyyy") ' Format(Left(Grid.TextMatrix(Cnt, 2), 2) & "-" & Mid(Grid.TextMatrix(Cnt, 2), 3, 2) & "-" & Right(Grid.TextMatrix(Cnt, 2), 4), "dd Mmm yyyy")

        rs1.Fields("ChequeAmt").Value = Grid.TextMatrix(Cnt, 3)
        rs1.Fields("RcptCode").Value = Grid.TextMatrix(Cnt, 13)
        rs1.Fields("SBEntryCode").Value = Grid.TextMatrix(Cnt, 16)

        'For DD Generated from Cash Deposit Slip
        If Left(Grid.TextMatrix(Cnt, 13), 2) = "CS" Then
            'Find that Cash Deposit Slip and change
            rs2.Open "Update DepositSlip set DDToCMSFlag = 1, DDToCMSSlipCode = " & txtCode.Text & " where SlipCode = " & Right(Grid.TextMatrix(Cnt, 13), Len(Grid.TextMatrix(Cnt, 13)) - 2) & " and CashFlag =1 and CashToOptFlag = 1", cn1, adOpenDynamic, adLockPessimistic
        End If


        rs1.Update
    End If
Next
rs1.Close

cn1.Close
End Sub

Public Sub CmdShow_Click()
On Error Resume Next
If DateDiff("d", Date, Me.dtSlip.Value) > 1 Then
    MsgBox "Slip date can't be more than " & Format(DateAdd("d", 1, Date), "dd Mmm yyyy") & ".", vbCritical + vbOKOnly
    dtSlip.SetFocus
    Exit Sub
End If

FrmCMSDeposit_Sub1.PrevCnt = 1
FrmCMSDeposit_Sub1.PrevVal = ""
Grid.Rows = 1
FillDepositSlipDetails
If cmdSave.Enabled = True Then FillPendingSlipDetails
FindTotals
End Sub

Private Sub FillDepositSlipDetails()
On Error Resume Next
If txtCode.Text = "" Then Exit Sub

Dim cnSys As ADODB.Connection
Set cnSys = New ADODB.Connection

Dim cnYr As ADODB.Connection
Set cnYr = New ADODB.Connection

Dim CnAim As ADODB.Connection
Set CnAim = New ADODB.Connection

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

Dim rs3 As ADODB.Recordset
Set rs3 = New ADODB.Recordset

Dim rs4 As ADODB.Recordset
Set rs4 = New ADODB.Recordset


Dim rs5 As ADODB.Recordset
Set rs5 = New ADODB.Recordset


CnAim.Open ModInit.ConnectStringAIM, , "panaaim"

Dim fso As FileSystemObject
Set fso = New FileSystemObject

rs1.Open "Select * from DepositSlip  where SlipCode =" & txtCode.Text, CnAim, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    If Val("" & rs1!ClearingType) = 0 Then
        Me.OptLocal.Value = True
    Else
        Me.OptOut.Value = True
    End If
End If
rs1.Close

rs1.Open "Select * from DepositSlipDetails where SlipCode =" & txtCode.Text & " order by entrycode", CnAim, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        Grid.Rows = Grid.Rows + 1
        Grid.TextMatrix(Grid.Rows - 1, 0) = ChqNewStatus
        Grid.TextMatrix(Grid.Rows - 1, 1) = rs1!ChequeNo
        Grid.TextMatrix(Grid.Rows - 1, 2) = Format(rs1.Fields("ChequeDate").Value, "ddMmyyyy")
        Grid.TextMatrix(Grid.Rows - 1, 3) = Format(rs1.Fields("ChequeAmt").Value, "0.00")
        Grid.TextMatrix(Grid.Rows - 1, 4) = "002"
        Grid.TextMatrix(Grid.Rows - 1, 5) = "002"
        Grid.TextMatrix(Grid.Rows - 1, 6) = "10"

        Grid.TextMatrix(Grid.Rows - 1, 10) = rs1!YearName
        Grid.TextMatrix(Grid.Rows - 1, 11) = rs1!BatchCode
        Grid.TextMatrix(Grid.Rows - 1, 12) = rs1!RollNo
        Grid.TextMatrix(Grid.Rows - 1, 13) = rs1!RcptCode

        Grid.TextMatrix(Grid.Rows - 1, 14) = Format(rs1!ChequeDate, "dd MMM yyyy")

        'Open yearfile to read other details
        'Check for yearname
        YrName = rs1!YearName

        'Open the year file
        YrPath = App.Path & "\smf" & StrReverse(Left(YrName, 4)) & "yr1.scd"

        If Left(rs1!RcptCode, 2) = "CS" Then
            Grid.TextMatrix(Grid.Rows - 1, 7) = "Cash Collection to DD"
        End If

        If fso.FileExists(YrPath) Then
            cnYr.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;User ID=admin;Extended Properties=""DBQ=" & YrPath & ";DefaultDir= " & App.Path & ";Driver={Microsoft Access Driver (*.mdb)};DriverId=281;FIL=MS Access;FILEDSN=" & YrPath & ";MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"";Initial Catalog=" & YrPath & """"
            cnYr.Open , , "panatech"

            If Left(rs1!RcptCode, 1) = "O" Then
                rs2.Open "Select * from StudentOtherPayment where (ChkNo ='" & rs1!ChequeNo & "' or ChkNo ='D" & rs1!ChequeNo & "') and RcptCode ='" & Replace(rs1!RcptCode, "O", "") & "'", cnYr, adOpenDynamic, adLockReadOnly
            Else
                rs2.Open "Select * from StudentPayment where  RecordDelFlag = 0 and (ChkNo ='" & rs1!ChequeNo & "' or ChkNo ='D" & rs1!ChequeNo & "') and RcptCode ='" & rs1!RcptCode & "'", cnYr, adOpenDynamic, adLockReadOnly
            End If

            If Not (rs2.BOF And rs2.EOF) Then
                If rs2.Fields("ChkStatus").Value = "Received at CO" Then
                    Grid.TextMatrix(Grid.Rows - 1, 0) = "Received at CO"
                ElseIf rs2.Fields("ChkStatus").Value = "Returned to Center" Then
                    Grid.TextMatrix(Grid.Rows - 1, 0) = "Returned to Center"
                ElseIf rs2.Fields("ChkStatus").Value = "Deposited" Then
                    Grid.TextMatrix(Grid.Rows - 1, 0) = ChqNewStatus
                End If

                rs3.Open "Select Title, FirstName, MidName, LastName, Sex, SBEntryCode from Student, StudentBatch where StudentBatch.RecordDelFlag = 0 and Student.StudentCode = StudentBatch.StudentCode and StudentBatch.BatchCode ='" & rs2.Fields("BatchCode").Value & "' and StudentBatch.RollNo =" & rs2.Fields("RollNo").Value, cnYr, adOpenDynamic, adLockReadOnly
                If Not (rs3.BOF And rs3.EOF) Then
                    Grid.TextMatrix(Grid.Rows - 1, 7) = rs3.Fields("FirstName").Value & " " & Left(rs3.Fields("MidName").Value, 1) & " " & rs3.Fields("LastName").Value
                    Grid.TextMatrix(Grid.Rows - 1, 9) = rs3!Sex
                    Grid.TextMatrix(Grid.Rows - 1, 16) = rs3!SBEntryCode
                End If
                rs3.Close

                Grid.TextMatrix(Grid.Rows - 1, 15) = "" & rs2!BankName

                rs3.Open "SELECT Streams.StreamName FROM StudentBatch INNER JOIN (Batches INNER JOIN Streams ON Batches.StreamCode = Streams.StreamCode) ON StudentBatch.BatchCode = Batches.BatchCode WHERE  StudentBatch.RecordDelFlag = 0 and StudentBatch.BatchCode='" & rs2.Fields("BatchCode").Value & "' AND StudentBatch.RollNo=" & rs2.Fields("RollNo").Value, cnYr, adOpenDynamic, adLockReadOnly
                If Not (rs3.BOF And rs3.EOF) Then
                    Grid.TextMatrix(Grid.Rows - 1, 8) = rs3!StreamName
                End If
                rs3.Close
            End If
            rs2.Close
            cnYr.Close
        End If

        rs1.MoveNext
    Loop
End If
rs1.Close
CnAim.Close

End Sub

Private Sub FillPendingSlipDetails()
On Error Resume Next
Dim cnSys As ADODB.Connection
Set cnSys = New ADODB.Connection

Dim cnYr As ADODB.Connection
Set cnYr = New ADODB.Connection

Dim CnAim As ADODB.Connection
Set CnAim = New ADODB.Connection

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

Dim rs3 As ADODB.Recordset
Set rs3 = New ADODB.Recordset

Dim rs4 As ADODB.Recordset
Set rs4 = New ADODB.Recordset

cnSys.ConnectionString = ModInit.ConnectStringSystem
cnSys.Open , , "panatech"

CnAim.Open ModInit.ConnectStringAIM, , "panaaim"

rs1.Open "Alter table DepositSlip add column CashFlag integer Null, CashToOptFlag integer Null, DDNo text(50) Null, DDDate DateTime Null, DDBankName Text(50) Null, CODepositDate DateTime Null, COSentThrough Text(50) Null, COConfirmFlag integer Null, DDToCMSFlag integer Null, DDToCMSSlipCode integer Null", CnAim, adOpenDynamic, adLockPessimistic

Dim fso As FileSystemObject
Set fso = New FileSystemObject

Dim AcceptEntryFlag As Boolean

rs1.Open "Select * from YearInfo order by DOLA desc", cnSys, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        YrPath = App.Path & "\" & rs1.Fields("FileName").Value
        YrName = rs1.Fields("Year").Value

        If fso.FileExists(YrPath) Then
            cnYr.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;User ID=admin;Extended Properties=""DBQ=" & YrPath & ";DefaultDir= " & App.Path & ";Driver={Microsoft Access Driver (*.mdb)};DriverId=281;FIL=MS Access;FILEDSN=" & YrPath & ";MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"";Initial Catalog=" & YrPath & """"
            cnYr.Open , , "panatech"

            rs2.Open "CREATE TABLE StudentOtherPayment (RcptCode Text (20)  NULL, BatchCode Text (15)  NULL, RollNo int NULL, AmountPaid Text (50)  NULL , AmtInWords Text (50)  NULL , RcptNo Text (15)  NULL , ChkNo Text (20)  NULL , BankName Text (50)  NULL, ChkDate DateTime NULL, NextPayDate DateTime NULL, PayDate DateTime NULL, ChkStatus Text (50)  NULL, RealDate DateTime NULL, Tip Text (50)  NULL, ModFlag Text (50)  NULL, OldRcptNo Text (15)  NULL, ModuleFlag Text (50)  NULL, StaxPer Text (50)  NULL, STaxAmt Text (50)  NULL, DepositDate DateTime NULL, UploadFlag Text (50)  NULL, AmountPaidReal real NULL, DelFlag Int) ", cnYr, adOpenDynamic, adLockPessimistic

            Dim ChkDateSerStr As String
            If Me.lblDepType.Caption = "CO" Then
                ChkDateSerStr = ""
            Else
                ChkDateSerStr = " and ChkDate <=#" & Format(dtSlip.Value, "dd Mmm yyyy") & "# "
            End If

            Dim OrderBySerStr As String
            If Me.lblDepType.Caption = "CO" Then
                OrderBySerStr = "order by SBEntryCode, chkdate"
            Else
                OrderBySerStr = "order by chkdate"
            End If

            rs2.Open "Select * from StudentPayment where RecordDelFlag = 0 and (ChkStatus = 'Pending' or ChkStatus = 'Returned to Center')  and ChkNo <> '' and chkno <> '000000'  and chkdate >#30 Sep 2008# " & ChkDateSerStr & " and RcptNo <> 'VE' " & OrderBySerStr, cnYr, adOpenDynamic, adLockReadOnly
            If Not (rs2.BOF And rs2.EOF) Then
                rs2.MoveFirst
                Do While Not rs2.EOF
                    If "" & rs2.Fields("ChkNo").Value <> "" Then
                        AcceptEntryFlag = False

'                        If rs2!chkNo = "628726" Then
'                            MsgBox ""
'                        End If

                        rs3.Open "SELECT Streams.*, StudentBatch.ScenarioPendingFlag, StudentBatch.Status, StudentBatch.PendingFlag, StudentBatch.AdmnDate  FROM StudentBatch INNER JOIN (Batches INNER JOIN Streams ON Batches.StreamCode = Streams.StreamCode) ON StudentBatch.BatchCode = Batches.BatchCode WHERE  StudentBatch.RecordDelFlag = 0 and StudentBatch.BatchCode='" & rs2.Fields("BatchCode").Value & "' AND StudentBatch.RollNo=" & rs2.Fields("RollNo").Value & " and ((Status = True) or (Status=False and PendingFlag =1))", cnYr, adOpenDynamic, adLockReadOnly
                        If Not (rs3.BOF And rs3.EOF) Then
                            If rs3!MIS_InstituteCode = FrmCMSDeposit.cboInstCode.Text Then
                                If rs3!Status = True Then
                                    If Val("" & rs3.Fields("ScenarioPendingFlag").Value) > 0 Then
                                        AcceptEntryFlag = False
                                    Else

                                        AcceptEntryFlag = True
                                        If rs3!AdmnDate >= DateValue("1 Jul 2011") Then
                                            'Check if student has paid full amount
                                            Dim TotalFees, TotalPaid As Single
                                            TotalFees = 0
                                            TotalPaid = 0
                                            'Find Students Balance
                                            rs4.Open "Select TotalFees from StudentBatch where SBEntryCode ='" & rs2!SBEntryCode & "' and Status = True and RecordDelFlag =0", cnYr, adOpenDynamic, adLockReadOnly
                                            If Not (rs4.BOF And rs4.EOF) Then
                                                TotalFees = rs4!TotalFees
                                            End If
                                            rs4.Close

                                            rs4.Open "Select sum(AmountPaid) as TotalPaid from StudentPayment where SBEntryCode ='" & rs2!SBEntryCode & "' and RecordDelFlag =0", cnYr, adOpenDynamic, adLockReadOnly
                                            If Not (rs4.BOF And rs4.EOF) Then
                                                TotalPaid = rs4!TotalPaid
                                            End If
                                            rs4.Close

                                            If (TotalFees - TotalPaid) > 2000000 Then   'Balance more than 2000000 Rs
                                                AcceptEntryFlag = False
                                            End If

                                        Else
                                            AcceptEntryFlag = True
                                        End If
                                    End If
                                ElseIf rs3!Status = False And rs3!PendingFlag = 1 Then
                                    'Check if PendingChequeDepFlag for this stream is 1 or 0
                                    If Val("" & rs3!PendingChequeDepFlag) = 1 Then
                                        'Check no of cheques allowed in pending status for this stream
                                        Dim PendingChequeDepCnt As Integer
                                        PendingChequeDepCnt = Val("" & rs3!PendingChequeDepCnt)

                                        'Find no of cheques deposited for this student
                                        rs4.Open "Select count(*) as DepChqCnt from StudentPayment where SBEntryCode ='" & rs2!SBEntryCode & "' and RcptCode <> '" & rs2!RcptCode & "' and (ChkStatus ='Deposited' or ChkStatus ='Cleared')", cnYr, adOpenDynamic, adLockReadOnly
                                        If Not (rs4.BOF And rs4.EOF) Then
                                            If Val("" & rs4!DepChqCnt) >= PendingChequeDepCnt Then
                                                AcceptEntryFlag = False
                                            Else
                                                AcceptEntryFlag = True
                                            End If
                                        Else
                                            If PendingChequeDepCnt > 0 Then
                                                AcceptEntryFlag = True
                                            Else
                                                AcceptEntryFlag = False
                                            End If
                                        End If
                                        rs4.Close
                                    Else
                                        AcceptEntryFlag = False
                                    End If
                                End If
                            Else
                                AcceptEntryFlag = False
                            End If
                        Else
                            AcceptEntryFlag = False
                        End If
                        rs3.Close



                        If AcceptEntryFlag = True Then
                            If rs2!chkstatus = "Pending" Then
                                AcceptEntryFlag = False
                                'Check if this rcpt is already included in any other deposit slip or no
                                rs4.Open "Select * from DepositSlipDetails where YearName ='" & YrName & "' and RcptCode ='" & rs2!RcptCode & "'", CnAim, adOpenDynamic, adLockReadOnly
                                If (rs4.BOF And rs4.EOF) Then
                                    AcceptEntryFlag = True
                                End If
                                rs4.Close
                            Else
                                AcceptEntryFlag = True
                            End If

                            If AcceptEntryFlag = True Then
                                Grid.Rows = Grid.Rows + 1
                                Grid.TextMatrix(Grid.Rows - 1, 1) = Replace(rs2!chkNo, "D", "")
                                If Me.lblDepType.Caption = "CO" Then
                                    Grid.TextMatrix(Grid.Rows - 1, 2) = Format(rs2.Fields("ChkDate").Value, "dd Mmm yyyy")
                                Else
                                    Grid.TextMatrix(Grid.Rows - 1, 2) = Format(rs2.Fields("ChkDate").Value, "ddMmyyyy")
                                End If
                                Grid.TextMatrix(Grid.Rows - 1, 3) = Format(rs2.Fields("AmountPaid").Value, "0.00")
                                Grid.TextMatrix(Grid.Rows - 1, 4) = "002"
                                Grid.TextMatrix(Grid.Rows - 1, 5) = "002"
                                Grid.TextMatrix(Grid.Rows - 1, 6) = "10"

                                StudentName = ""
                                rs3.Open "Select Title, FirstName, MidName, LastName, Sex, SBEntryCode from Student, StudentBatch where Student.StudentCode = StudentBatch.StudentCode and StudentBatch.RecordDelFlag =0 and StudentBatch.BatchCode ='" & rs2.Fields("BatchCode").Value & "' and StudentBatch.RollNo =" & rs2.Fields("RollNo").Value, cnYr, adOpenDynamic, adLockReadOnly
                                If Not (rs3.BOF And rs3.EOF) Then
                                    Grid.TextMatrix(Grid.Rows - 1, 7) = rs3.Fields("FirstName").Value & " " & Left(rs3.Fields("MidName").Value, 1) & " " & rs3.Fields("LastName").Value
                                    Grid.TextMatrix(Grid.Rows - 1, 9) = rs3!Sex
                                    Grid.TextMatrix(Grid.Rows - 1, 16) = rs3!SBEntryCode

                                End If
                                rs3.Close

    '                            If Trim(Grid.TextMatrix(Grid.Rows - 1, 7)) = "" Then MsgBox ""

                                rs3.Open "SELECT Streams.StreamName FROM StudentBatch INNER JOIN (Batches INNER JOIN Streams ON Batches.StreamCode = Streams.StreamCode) ON StudentBatch.BatchCode = Batches.BatchCode WHERE  StudentBatch.RecordDelFlag = 0 and StudentBatch.BatchCode='" & rs2.Fields("BatchCode").Value & "' AND StudentBatch.RollNo=" & rs2.Fields("RollNo").Value, cnYr, adOpenDynamic, adLockReadOnly
                                If Not (rs3.BOF And rs3.EOF) Then
                                    Grid.TextMatrix(Grid.Rows - 1, 8) = rs3!StreamName
                                End If
                                rs3.Close

                                Grid.TextMatrix(Grid.Rows - 1, 10) = YrName
                                Grid.TextMatrix(Grid.Rows - 1, 11) = rs2!BatchCode
                                Grid.TextMatrix(Grid.Rows - 1, 12) = rs2!RollNo
                                Grid.TextMatrix(Grid.Rows - 1, 13) = rs2!RcptCode

                                Grid.TextMatrix(Grid.Rows - 1, 14) = Format(rs2!ChkDate, "dd MMM yyyy")
                                Grid.TextMatrix(Grid.Rows - 1, 15) = "" & rs2!BankName
                            End If
                        End If
                    End If
                    rs2.MoveNext
                Loop
            End If
            rs2.Close

            'Read Otherpayment data
            rs2.Open "Select * from StudentOtherPayment where (ChkStatus = 'Pending' or ChkStatus = 'Returned to Center') and ChkNo <> '' and chkdate >#30 Sep 2008# and ChkDate <=#" & Format(dtSlip.Value, "dd Mmm yyyy") & "# and RcptNo <> 'VE' order by chkdate", cnYr, adOpenDynamic, adLockReadOnly
            If Not (rs2.BOF And rs2.EOF) Then
                rs2.MoveFirst
                Do While Not rs2.EOF
                    If "" & rs2.Fields("ChkNo").Value <> "" Then

                        'Check if this rcpt is already included in any other deposit slip or no
                        rs4.Open "Select * from DepositSlipDetails where YearName ='" & YrName & "' and RcptCode ='O" & rs2!RcptCode & "'", CnAim, adOpenDynamic, adLockReadOnly
                        If (rs4.BOF And rs4.EOF) Then
                            Grid.Rows = Grid.Rows + 1
                            Grid.TextMatrix(Grid.Rows - 1, 1) = Replace(rs2!chkNo, "D", "")
                            Grid.TextMatrix(Grid.Rows - 1, 2) = Format(rs2.Fields("ChkDate").Value, "ddMmyyyy")
                            Grid.TextMatrix(Grid.Rows - 1, 3) = Format(rs2.Fields("AmountPaid").Value, "0.00")
                            Grid.TextMatrix(Grid.Rows - 1, 4) = "002"
                            Grid.TextMatrix(Grid.Rows - 1, 5) = "002"
                            Grid.TextMatrix(Grid.Rows - 1, 6) = "10"

                            StudentName = ""
                            rs3.Open "Select Title, FirstName, MidName, LastName, Sex from Student, StudentBatch where StudentBatch.RecordDelFlag =0 and Student.StudentCode = StudentBatch.StudentCode and StudentBatch.BatchCode ='" & rs2.Fields("BatchCode").Value & "' and StudentBatch.RollNo =" & rs2.Fields("RollNo").Value, cnYr, adOpenDynamic, adLockReadOnly
                            If Not (rs3.BOF And rs3.EOF) Then
                                Grid.TextMatrix(Grid.Rows - 1, 7) = rs3.Fields("FirstName").Value & " " & Left(rs3.Fields("MidName").Value, 1) & " " & rs3.Fields("LastName").Value
                                Grid.TextMatrix(Grid.Rows - 1, 9) = rs3!Sex
                            End If
                            rs3.Close


                            rs3.Open "SELECT Streams.StreamName FROM StudentBatch INNER JOIN (Batches INNER JOIN Streams ON Batches.StreamCode = Streams.StreamCode) ON StudentBatch.BatchCode = Batches.BatchCode WHERE StudentBatch.RecordDelFlag = 0 and StudentBatch.BatchCode='" & rs2.Fields("BatchCode").Value & "' AND StudentBatch.RollNo=" & rs2.Fields("RollNo").Value, cnYr, adOpenDynamic, adLockReadOnly
                            If Not (rs3.BOF And rs3.EOF) Then
                                Grid.TextMatrix(Grid.Rows - 1, 8) = rs3!StreamName
                            End If
                            rs3.Close

                            Grid.TextMatrix(Grid.Rows - 1, 10) = YrName
                            Grid.TextMatrix(Grid.Rows - 1, 11) = rs2!BatchCode
                            Grid.TextMatrix(Grid.Rows - 1, 12) = rs2!RollNo
                            Grid.TextMatrix(Grid.Rows - 1, 13) = "O" & rs2!RcptCode

                            Grid.TextMatrix(Grid.Rows - 1, 14) = Format(rs2!ChkDate, "dd MMM yyyy")
                            Grid.TextMatrix(Grid.Rows - 1, 15) = "" & rs2!BankName
                        End If
                        rs4.Close

                    End If
                    rs2.MoveNext
                Loop
            End If
            rs2.Close
            cnYr.Close
        End If
        rs1.MoveNext
    Loop
End If
rs1.Close

'Add entries from Cash Deposit Table
rs1.Open "Select * from DepositSlip where CashFlag =1 and DDToCMSFlag= 0 and CashToOptFlag =1 and DDdate >#30 Sep 2008# and DDDate <=#" & Format(dtSlip.Value, "dd Mmm yyyy") & "# order by DDDate", CnAim, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        Grid.Rows = Grid.Rows + 1
        Grid.TextMatrix(Grid.Rows - 1, 1) = Replace(rs1!DDNo, "D", "")
        Grid.TextMatrix(Grid.Rows - 1, 2) = Format(rs1.Fields("DDDate").Value, "ddMmyyyy")
        Grid.TextMatrix(Grid.Rows - 1, 3) = Format(rs1.Fields("ChqAmt").Value, "0.00")
        Grid.TextMatrix(Grid.Rows - 1, 4) = "002"
        Grid.TextMatrix(Grid.Rows - 1, 5) = "002"
        Grid.TextMatrix(Grid.Rows - 1, 6) = "10"

        Grid.TextMatrix(Grid.Rows - 1, 7) = "Cash Collection to DD"
        Grid.TextMatrix(Grid.Rows - 1, 9) = ""


        Grid.TextMatrix(Grid.Rows - 1, 8) = rs1!SlipNo ' rs3!StreamName

        Grid.TextMatrix(Grid.Rows - 1, 10) = ""  'YrName
        Grid.TextMatrix(Grid.Rows - 1, 11) = ""  'rs2!BatchCode
        Grid.TextMatrix(Grid.Rows - 1, 12) = ""  'rs2!RollNo
        Grid.TextMatrix(Grid.Rows - 1, 13) = "CS" & rs1!SlipCode

        Grid.TextMatrix(Grid.Rows - 1, 14) = Format(rs1!DDDate, "dd MMM yyyy")
        Grid.TextMatrix(Grid.Rows - 1, 15) = "" & rs1!DDBankName

        rs1.MoveNext
    Loop
End If
rs1.Close

cnSys.Close
End Sub

Private Sub cmdSkip_Click()
On Error Resume Next
If Trim(txtSlipNo.Text) = "" Then Exit Sub
If MsgBox("You are about to Skip a CMS Deposit Slip No.It is an Irreversible Process.To confirm Click 'Yes'", vbCritical + vbYesNo) = vbNo Then Exit Sub

Dim rs1 As New ADODB.Recordset
Dim cn1 As New ADODB.Connection

cn1.Open ConnectStringSystem, , "panatech"

rs1.Open "Select* from SkipDepositSlips where SeriesCode = " & TxtSeriesCode.Text & " and DepositSlipNo ='" & txtSlipNo.Text & "'", cn1, adOpenDynamic, adLockPessimistic
rs1.AddNew
rs1.Fields("SeriesCode") = Val(TxtSeriesCode.Text)
rs1.Fields("DepositSlipNo") = Val(txtSlipNo.Text)
rs1.Update
rs1.Close
cn1.Close
AutoDepositSlipNo
End Sub

Private Sub dtSlip_Change()
On Error Resume Next
If ModInit.HTGroup = True Then
    AutoDepositSlipNoHT
Else
    AutoDepositSlipNo
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
'If Grid.Rows = 1 Then Exit Sub
'Me.PrevCnt = 1
AdmnFlag = False
If KeyCode = vbKeyF3 Then
    FrmCMSDeposit_Sub2.Show vbModal
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
'ModInit.Init
ModInit.SetFormColor Me, "Orange"
ModInit.SetGridColor Grid, "Orange"
AdmnFlag = False
CmsChkFlag = True
Grid.ToolTipText = "Double click on any Cheque to select it for Deposit"
lblDepType.FontSize = 24

Shape2.FillColor = ModInit.SubColor
Shape4.FillColor = ModInit.SubColor

OptLocal.BackColor = ModInit.SubColor
OptOut.BackColor = ModInit.SubColor
ChqNewStatus = "Deposit"

txtFlag.Text = "ADD"
ModFlag = False
dtSlip.Value = Date
OptLocal.Value = True
With Grid
    .WordWrap = True
    .RowHeight(0) = 600
    .ColWidth(0) = 1500
    .ColWidth(1) = 1000
    .ColWidth(2) = 1100
    .ColWidth(3) = 1000
    .ColWidth(4) = 500
    .ColWidth(5) = 500
    .ColWidth(6) = 500
    .ColWidth(7) = 2500
    .ColWidth(8) = 2000
    .ColWidth(9) = 500
    .ColWidth(10) = 1000
    .ColWidth(11) = 0
    .ColWidth(12) = 0
    .ColWidth(15) = 0
    .ColWidth(13) = 0
    .ColWidth(16) = 0
    .Rows = 1
End With
PrevCnt = 1
PrevVal = ""
If ModInit.HTGroup = True Then
    txtSlipNo.Locked = False
    AutoDepositSlipNoHT
'    cmdSkip.Visible = False
Else
'    cmdSkip.Visible = True
    txtSlipNo.Locked = True
    AutoDepositSlipNo
End If
End Sub

Private Sub AutoDepositSlipNo()
On Error Resume Next
Dim rs1 As New ADODB.Recordset
Dim cnSys As New ADODB.Connection
Dim CnAim As New ADODB.Connection
Dim rs2 As New ADODB.Recordset

Dim DepositSlipNoPre, DepositSlipNoPost

cnSys.Open ConnectStringSystem, , "panatech"
CnAim.Open ConnectStringAIM, , "panaaim"

rs1.Open "Select * from CM_YearMaster where FromDate <=#" & dtSlip.Value & "# and Todate >=#" & dtSlip.Value & "#", cnSys, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    DepositSlipNoPre = FrmCMSDeposit.cboInstCode.Text & ModInit.Liccode & Val(rs1.Fields("Cmsyear").Value)
End If
rs1.Close

If DepositSlipNoPre = "" Then
    If dtSlip.Value <= DateValue("31 Mar " & Year(Date)) Then
        DepositSlipNoPre = FrmCMSDeposit.cboInstCode.Text & ModInit.Liccode & Val(Right(Year(Date), 1))
    Else
        DepositSlipNoPre = FrmCMSDeposit.cboInstCode.Text & ModInit.Liccode & (Val(Right(Year(Date), 1)) + 1)
    End If
End If

If Trim(DepositSlipNoPre) <> "" Then
    rs1.Open "Select max(right(SlipNo,4)) As MainVal from DepositSlip where Left(SlipNo,6)='" & DepositSlipNoPre & "'", CnAim, adOpenDynamic, adLockReadOnly
    If Not (rs1.BOF And rs1.EOF) Then
        DepositSlipNoPost = Val("" & rs1.Fields("MainVal").Value)
    End If
    rs1.Close

    txtSlipNo.Text = DepositSlipNoPre & Format(Val(DepositSlipNoPost) + 1, "0000")
Else
    txtSlipNo.Text = ""
End If

cnSys.Close
CnAim.Close
End Sub

Public Sub AutoDepositSlipNoHT()
On Error Resume Next
Dim rs1 As New ADODB.Recordset
Dim cnSys As New ADODB.Connection
Dim CnAim As New ADODB.Connection
Dim rs2 As New ADODB.Recordset

Dim DepositSlipNoPre, DepositSlipNoPost

cnSys.Open ConnectStringSystem, , "panatech"
CnAim.Open ConnectStringAIM, , "panaaim"


DepositSlipNoPre = FrmCMSDeposit.cboInstCode.Text & ModInit.Liccode & GetCMSYearCode(Format(dtSlip.Value, "YYYY"))

If lblDepType.Caption = "Bank" Then
    DepositSlipNoPre = DepositSlipNoPre & "B"
Else
    DepositSlipNoPre = DepositSlipNoPre & "C"
End If

If Trim(DepositSlipNoPre) <> "" Then
    rs1.Open "Select max(right(SlipNo,3)) As MainVal from DepositSlip where DepositSlipType ='" & lblDepType.Caption & "' and Left(SlipNo,7)='" & DepositSlipNoPre & "'", CnAim, adOpenDynamic, adLockReadOnly
    If Not (rs1.BOF And rs1.EOF) Then
        DepositSlipNoPost = Val("" & rs1.Fields("MainVal").Value)
    End If
    rs1.Close

    txtSlipNo.Text = DepositSlipNoPre & Format(Val(DepositSlipNoPost) + 1, "000")
Else
    txtSlipNo.Text = ""
End If

cnSys.Close
CnAim.Close
End Sub

Private Function GetCMSYearCode(YearName As String) As String
On Error Resume Next
Select Case YearName
    Case "2011": GetCMSYearCode = "A"
    Case "2012": GetCMSYearCode = "B"
    Case "2013": GetCMSYearCode = "C"
    Case "2014": GetCMSYearCode = "D"
    Case "2015": GetCMSYearCode = "E"
    Case "2016": GetCMSYearCode = "F"
    Case "2017": GetCMSYearCode = "G"
    Case "2018": GetCMSYearCode = "H"
    Case "2019": GetCMSYearCode = "I"
    Case "2020": GetCMSYearCode = "J"
    Case "2021": GetCMSYearCode = "L"
    Case "2022": GetCMSYearCode = "L"
    Case "2023": GetCMSYearCode = "M"
    Case "2024": GetCMSYearCode = "N"
    Case "2025": GetCMSYearCode = "O"
End Select
End Function

Private Sub AutoDepositSlipNoOld()
On Error Resume Next
Dim rs1 As New ADODB.Recordset
Dim cnSys As New ADODB.Connection
Dim CnAim As New ADODB.Connection
Dim rs2 As New ADODB.Recordset

Dim DepositSlipNo, Ser_Code

cnSys.Open ConnectStringSystem, , "panatech"
CnAim.Open ConnectStringAIM, , "panaaim"

rs1.Open "Select * from DepositSlipSeries where Approval=1 order by SeriesCode desc ", cnSys, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    TxtSeriesCode.Text = rs1.Fields("SeriesCode")
    rs2.Open "Select Max(SlipNo) from DepositSlip where SeriesCode=" & rs1.Fields("SeriesCode"), CnAim, adOpenDynamic, adLockReadOnly
    If Val("" & rs2.Fields(0)) = 0 Then
        DepositSlipNo = rs1.Fields("SeriesFrom")
    Else
        DepositSlipNo = rs2.Fields(0) + 1
    End If
    rs2.Close

    Do While RecFoundFlag = False
        rs2.Open "Select * from SkipDepositSlips where SeriesCode =" & rs1.Fields("SeriesCode") & " and DepositSlipNo= '" & DepositSlipNo & "'", cnSys, adOpenDynamic, adLockPessimistic
        If (rs2.BOF And rs2.EOF) Then
            RecFoundFlag = True
        Else
            DepositSlipNo = DepositSlipNo + 1
        End If
        rs2.Close
    Loop
    Ser_Code = Val(rs1.Fields("SeriesCode"))
End If
rs1.Close

rs2.Open "Select * from DepositSlipSeries where SeriesCode=" & Ser_Code & "  and SeriesFrom <='" & DepositSlipNo & "' and  SeriesTo>= '" & DepositSlipNo & "'", cnSys, adOpenDynamic, adLockReadOnly
If Not (rs2.BOF And rs2.EOF) Then
    txtSlipNo.Text = DepositSlipNo
End If
rs2.Close
cnSys.Close
CnAim.Close
End Sub

Private Function checkvalid() As Boolean
On Error Resume Next
If Trim(Me.txtSlipNo.Text) = "" Then
    MsgBox "Enter Deposit Slip Number.", vbInformation + vbOKOnly
    txtSlipNo.SetFocus
    checkvalid = False
    Exit Function
End If

If Val(Me.txtChqCnt.Text) <= 0 Then
    MsgBox "You have not selected any Cheque for this Deposit Slip.  Select Cheque and then Save the entry.", vbCritical + vbOKOnly
    Grid.SetFocus
    checkvalid = False
    Exit Function
End If

checkvalid = True
End Function

Private Sub ReadNewCode()
On Error Resume Next
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection
cn1.ConnectionString = ModInit.ConnectStringAIM
cn1.Open , , "panaaim"

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset
rs1.Open "select max(SlipCode) as LastSlipCode from DepositSlip", cn1, adOpenDynamic, adLockOptimistic
If Not (rs1.BOF And rs1.EOF) Then
    SlipId = Val("" & rs1.Fields("LastSlipCode")) + 1
Else
    SlipId = 1
End If
rs1.Close
cn1.Close
End Sub

Private Sub Grid_DblClick()
On Error Resume Next
If Grid.Rows = 1 Then Exit Sub

If AdmnFlag = False Then
    If ModInit.UserType <> 0 Then Exit Sub
End If

If cmdSave.Enabled = False Then
'    If ModInit.UserType = 0 Then
'        res = MsgBox("You are about to remove selected cheque from Finalised CMS Deposit Slip.  System will store this action in database for future reference.  Do you want to proceed?", vbQuestion + vbYesNo)
'        If res = 7 Then Exit Sub
'
'        Dim cn1 As ADODB.Connection
'        Set cn1 = New ADODB.Connection
'
'        cn1.ConnectionString = ModInit.ConnectStringAIM
'        cn1.Open , , "panaaim"
'
'        Dim rs1 As ADODB.Recordset
'        Set rs1 = New ADODB.Recordset
'
'        Dim rs2 As ADODB.Recordset
'        Set rs2 = New ADODB.Recordset
'
'        Dim cnYr As New ADODB.Connection
'
'        Dim fso As FileSystemObject
'        Set fso = New FileSystemObject
'
'        'Add entry in database
'        rs1.Open "Select * from DepositSlipDetails where SlipCode =" & txtCode.Text & " and ChequeNo = '" & Grid.TextMatrix(Grid.RowSel, 1) & "' and RcptCode ='" & Grid.TextMatrix(Grid.RowSel, 13) & "'", cn1, adOpenDynamic, adLockReadOnly
'        If Not (rs1.BOF And rs1.EOF) Then
'            rs2.Open "Select * from DepositSlipDetailsDel where SlipCode =" & txtCode.Text & " and ChequeNo = '" & Grid.TextMatrix(Grid.RowSel, 1) & "' and RcptCode ='" & Grid.TextMatrix(Grid.RowSel, 13) & "'", cn1, adOpenDynamic, adLockPessimistic
'            rs2.AddNew
'            rs2!SlipCode = rs1!SlipCode
'            rs2!EntryCode = rs1!EntryCode
'            rs2!YearName = rs1!YearName
'            rs2!BatchCode = rs1!BatchCode
'            rs2!RollNo = rs1!RollNo
'            rs2!ChequeNo = rs1!ChequeNo
'            rs2!ChequeDate = rs1!ChequeDate
'            rs2!ChequeAmt = rs1!ChequeAmt
'            rs2!RcptCode = rs1!RcptCode
'            rs2!DelDate = Date
'            rs2.Update
'            rs2.Close
'        End If
'        rs1.Close
'
'
'        'Delete cheque from CMS Deposit Slip
'        rs1.Open "Delete from DepositSlipDetails where SlipCode =" & txtCode.Text & " and ChequeNo = '" & Grid.TextMatrix(Grid.RowSel, 1) & "' and RcptCode ='" & Grid.TextMatrix(Grid.RowSel, 13) & "'", cn1, adOpenDynamic, adLockPessimistic
'
'        'Mark cheque as Pending
'        YrName = Grid.TextMatrix(Grid.RowSel, 10)
'
'        'Open the year file
'        YrPath = App.Path & "\smf" & StrReverse(Left(YrName, 4)) & "yr1.scd"
'
'        If fso.FileExists(YrPath) Then
'            cnYr.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;User ID=admin;Extended Properties=""DBQ=" & YrPath & ";DefaultDir= " & App.Path & ";Driver={Microsoft Access Driver (*.mdb)};DriverId=281;FIL=MS Access;FILEDSN=" & YrPath & ";MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"";Initial Catalog=" & YrPath & """"
'            cnYr.Open , , "panatech"
'
'            If Left(Grid.TextMatrix(Grid.RowSel, 13), 1) = "O" Then
'                rs2.Open "Select * from StudentOtherPayment where ChkNo ='" & Grid.TextMatrix(Grid.RowSel, 1) & "' and RcptCode ='" & Replace(Grid.TextMatrix(Grid.RowSel, 13), "O", "") & "'", cnYr, adOpenDynamic, adLockPessimistic
'            Else
'                rs2.Open "Select * from StudentPayment where RecordDelFlag = 0 and ChkNo ='" & Grid.TextMatrix(Grid.RowSel, 1) & "' and RcptCode ='" & Grid.TextMatrix(Grid.RowSel, 13) & "'", cnYr, adOpenDynamic, adLockPessimistic
'            End If
'            If Not (rs2.BOF And rs2.EOF) Then
'                rs2.Fields("ChkStatus").Value = "Pending"
'                rs2.Fields("DepositDate").Value = rs2!PayDate
'                rs2.Fields("RealDate").Value = ""
'                rs2.Fields("SlipNo").Value = ""
'                rs2.Fields("UpLoadFlag").Value = "False"
'                rs2.Update
'            End If
'            rs2.Close
'
'            cnYr.Close
'        End If
'
'
'
'        'Remove cheque from grid
'        If Grid.Rows > 2 Then
'            Grid.RemoveItem (Grid.RowSel)
'        Else
'            Grid.Rows = 1
'        End If
'        FindTotals
'
'        'Update cheque cnt and amount total
'        rs1.Open "Select * from DepositSlip where SlipCode =" & txtCode.Text, cn1, adOpenDynamic, adLockPessimistic
'        If Not (rs1.BOF And rs1.EOF) Then
'            rs1.Fields("ChqCnt").Value = Val(Me.txtChqCnt.Text)
'            rs1.Fields("ChqAmt").Value = Val(Replace(Me.txtSlipAmt.Text, ",", ""))
'            rs1.Fields("UploadFlag").Value = "False"
'            rs1.Update
'        End If
'        rs1.Close
'        cn1.Close
'        Exit Sub
'    Else
        Exit Sub
'    End If
End If

If Grid.TextMatrix(Grid.RowSel, 0) = ChqNewStatus Then
    Grid.TextMatrix(Grid.RowSel, 0) = ""
Else
    'Check if total no of cheques is not more than 50
    Dim ChqCnt As Integer
    For Cnt = 1 To Grid.Rows - 1
        If Grid.TextMatrix(Cnt, 0) = ChqNewStatus Then
            ChqCnt = ChqCnt + 1
        End If
    Next
    If ChqCnt >= 50 Then
        MsgBox "You can't add more than 50 cheques in a Slip.  Save this Slip and then add remaining Cheques in another Slip.", vbInformation + vbOKOnly
        Exit Sub
    End If

    Grid.TextMatrix(Grid.RowSel, 0) = ChqNewStatus
End If
FindTotals
End Sub

Private Sub FindTotals()
On Error Resume Next
Dim ChqCnt As Integer
Dim ChqAmt As Double

For Cnt = 1 To Grid.Rows - 1
    If Grid.TextMatrix(Cnt, 0) = ChqNewStatus Then
        ChqCnt = ChqCnt + 1
        ChqAmt = ChqAmt + Val(Grid.TextMatrix(Cnt, 3))
    End If
Next

Me.txtChqCnt.Text = ChqCnt
Me.txtSlipAmt.Text = Format(ChqAmt, "#,##,##,##0.00")
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
On Error Resume Next
If CmsChkFlag = False Or Grid.TextMatrix((Grid.RowSel), 0) = ChqNewStatus Then
    AdmnFlag = True
    If KeyAscii = vbKeySpace Then
        Grid_DblClick
        KeyAscii = 0
    End If
    CmsChkFlag = True
End If
End Sub


