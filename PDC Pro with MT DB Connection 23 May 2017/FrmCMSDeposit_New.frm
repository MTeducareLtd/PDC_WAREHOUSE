VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FrmCMSDeposit_New 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CMS Slip Generator"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAdd2CMS 
      Caption         =   "&5 Add 2 CMS"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdShowNonMICR 
      Caption         =   "&4 Show"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdOutstation 
      Caption         =   "&3 Show"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmsShowNonICICI 
      Caption         =   "&2 Show"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   240
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdShowICICI 
      Caption         =   "&1 Show"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtSlip 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
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
      Format          =   125304835
      CurrentDate     =   39310
   End
   Begin SHDocVwCtl.WebBrowser wbSend 
      Height          =   735
      Left            =   1800
      TabIndex        =   16
      Top             =   3360
      Visible         =   0   'False
      Width           =   1815
      ExtentX         =   3201
      ExtentY         =   1296
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Non-MICR Cheques"
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
      Top             =   2520
      Width           =   1680
   End
   Begin VB.Label lblChqCnt_NonMICR 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
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
      Left            =   2880
      TabIndex        =   12
      Top             =   2520
      Width           =   135
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6240
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblChqCnt_Outstation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
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
      Left            =   2880
      TabIndex        =   11
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label lblChqCnt_NonICICI 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
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
      Left            =   2880
      TabIndex        =   10
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label lblChqCnt_ICICI 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
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
      Left            =   2880
      TabIndex        =   9
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Outstation Bank CMS"
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
      Top             =   2040
      Width           =   1830
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Local Non - ICICI Bank CMS"
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
      Top             =   1560
      Width           =   2430
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Local ICICI Bank CMS"
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
      Top             =   1080
      Width           =   1905
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CMS Slip Date"
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
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "FrmCMSDeposit_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdAdd2CMS_Click()
FrmCMSDeposit_New_Sub3.Show vbModal
End Sub

Private Sub cmdGenerate_Click()
On Error GoTo ErrExit
Me.lblChqCnt_ICICI.Caption = "0"
Me.lblChqCnt_NonICICI.Caption = "0"
Me.lblChqCnt_NonMICR.Caption = "0"
Me.lblChqCnt_Outstation.Caption = "0"

Me.cmdAdd2CMS.Visible = False

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

'Check if cms for this date is already generated or not
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

rs1.Open "Select * from ASPDC_CMS_NewLog where Location_Code ='" & ModInit.LocationCode & "' and CMS_Date ='" & Format(Me.dtSlip.Value, "dd Mmm yyyy") & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    'CMS already generated
    'Hence show print option
    
'    wbSend.Navigate2 "http://oe.mteducare.com/pdc_management/ImportStudentData_OrderEngine.aspx?DSC=" & Format(dtSlip.Value, "ddmmyyyy") & "&LOC=" & ModInit.LocationCode & ""
'    DoEvents
    
    FillCMSSummary
    
Else
    'Check is date is not less than current date
    If Me.dtSlip.Value < Date Then
        MsgBox "You can't generate CMS for date lesser than current date.", vbCritical + vbOKOnly
        dtSlip.SetFocus
        Exit Sub
    End If
    
    'Check if date is greater than current date + 5 days
    If Me.dtSlip.Value > DateAdd("d", 5, Date) Then
        MsgBox "You can't generate CMS for date greater than current date + 5 days.", vbCritical + vbOKOnly
        dtSlip.SetFocus
        Exit Sub
    End If
    
    
    Dim UserRes As Integer
    UserRes = MsgBox("You are about to generate CMS for date " & Format(dtSlip.Value, "dd Mmm yyyy") & " for location " & ModInit.LocationName & ".  Do you want to proceed now?", vbQuestion + vbYesNo)
    If UserRes = 6 Then
        'Generate CMS
        Dim cmd As ADODB.Command
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = cn1
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SP_ASPDC_CMSGenerateForICICIBank"
        
        cmd.Parameters.Append cmd.CreateParameter("CMSDate", adDate, adParamInput, , Me.dtSlip.Value)
        cmd.Parameters.Append cmd.CreateParameter("Location_Code", adVarChar, adParamInput, 50, ModInit.LocationCode)
        cmd.Parameters.Append cmd.CreateParameter("result", adInteger, adParamOutput)
        cmd.Execute
        
        res = cmd("result")
        
        If Val("" & res) >= 0 Then
            'Fill SBentryCodes for the students
            wbSend.Navigate2 "http://oe.mteducare.com/pdc_management/ImportStudentData_OrderEngine.aspx?DSC=" & Format(dtSlip.Value, "ddmmyyyy") & "&LOC=" & ModInit.LocationCode & ""
            DoEvents
        
        
            FillCMSSummary
            MsgBox "CMS generated successfully.", vbInformation + vbOKOnly
        Else
            MsgBox "Error in CMS generation.  Please try again or contact Acountech Solutions.", vbInformation + vbOKOnly
        End If
        
        Set cmd.ActiveConnection = Nothing
    

'        'Generate CMS
'        Dim cmd As ADODB.Command
'
'        Set cmd = New ADODB.Command
'        cmd.ActiveConnection = cn1
'        cmd.CommandType = adCmdStoredProc
'        cmd.CommandText = "SP_ASPDC_CMSGenerateForICICIBank"
'
'        cmd.Parameters.Append cmd.CreateParameter("CMSDate", adDate, adParamInput, Me.dtSlip.Value)
'        cmd.Parameters.Append cmd.CreateParameter("Location_Code", adVarChar, adParamInput, 10, ModInit.LocationCode)
'        cmd.Parameters.Append cmd.CreateParameter("result", adInteger, adParamOutput)
'
'        Dim res As Integer
'        cmd.Execute
'        res = cmd("result")
'
        
    End If
End If
rs1.Close
cn1.Close


Exit Sub
ErrExit:
MsgBox "Error: " & Err.Description, vbCritical + vbOKOnly

End Sub

Private Sub FillCMSSummary()
On Error Resume Next
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

'Check if cms for this date is already generated or not
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

Me.lblChqCnt_ICICI.Caption = "0"
Me.lblChqCnt_NonICICI.Caption = "0"
Me.lblChqCnt_NonMICR.Caption = "0"
Me.lblChqCnt_Outstation.Caption = "0"

If ModInit.LocationCode = "001" Then
    Me.cmdAdd2CMS.Visible = True
Else
    If ModInit.PDCUserName = "Mithun" Then
        Me.cmdAdd2CMS.Visible = True
    End If
End If

'ICICI Bank CMS
rs2.Open "select isnull(count(*),0) as RecCnt from ASPDC_DispatchSlipDetails where Location_Code ='" & ModInit.LocationCode & "' and cmsslipno ='" & Format(dtSlip.Value, "ddMMYYYY") & "' and right(left( MICRNumber,6),3) = '229'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs2.BOF And rs2.EOF) Then
    'Fill grid for this
    lblChqCnt_ICICI.Caption = rs2!RecCnt
End If
rs2.Close

'Non ICICI Bank CMS
rs2.Open "select isnull(count(*),0) as RecCnt from ASPDC_DispatchSlipDetails where Location_Code ='" & ModInit.LocationCode & "' and cmsslipno ='" & Format(dtSlip.Value, "ddMMYYYY") & "' and right(left( MICRNumber,6),3) <> '229' and left(MICRNumber,3) ='" & ModInit.MICRLocationCode & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs2.BOF And rs2.EOF) Then
    'Fill grid for this
    lblChqCnt_NonICICI.Caption = rs2!RecCnt
End If
rs2.Close

'Outstation Bank CMS
rs2.Open "select isnull(count(*),0) as RecCnt from ASPDC_DispatchSlipDetails where Location_Code ='" & ModInit.LocationCode & "' and cmsslipno ='" & Format(dtSlip.Value, "ddMMYYYY") & "' and right(left( MICRNumber,6),3) <> '229' and left(MICRNumber,3) <> '" & ModInit.MICRLocationCode & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs2.BOF And rs2.EOF) Then
    'Fill grid for this
    lblChqCnt_Outstation.Caption = rs2!RecCnt
End If
rs2.Close

'Non MICR
rs2.Open "select isnull(count(*),0) as RecCnt from ASPDC_DispatchSlipDetails where Location_Code ='" & ModInit.LocationCode & "' and cmsslipno ='" & Format(dtSlip.Value, "ddMMYYYY") & "' and (len(isnull(MICRNumber,''))) <> 9", cn1, adOpenDynamic, adLockReadOnly
If Not (rs2.BOF And rs2.EOF) Then
    'Fill grid for this
    lblChqCnt_NonMICR.Caption = rs2!RecCnt
End If
rs2.Close
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

Dim RecordCnt As Integer
RecordCnt = CMSGrid.Rows - 1

If RecordCnt = 0 Then Exit Sub

'For A4 size paper
LeftMargin = 200
RightMargin = 15000
TopMargin = 400
BottomMargin = 11000

PrintLines = 25

' Column headings and blank line
TROWS = CMSGrid.Rows - 1

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
    NoOfCopies = 4
'Else
'    NoOfCopies = 3
'End If

'Setup Global for printer
Printer.Font = "arial"
Printer.DrawWidth = 2
Printer.FillStyle = vbFSSolid
Printer.Orientation = cdlLandscape

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
            PrintMatter = "CITICLEAR DEPOSIT SLIP/DETAILS"
'        Else
'            PrintMatter = "CITICHECK DEPOSIT SLIP/DETAILS"
'        End If
        Printer.CurrentX = LeftMargin + 100
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        PrintMatter = "CITIBANK"
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
        PrintMatter = Grid.TextMatrix(RowNo, 3)
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
        PrintMatter = Format(dtSlip.Value, "dd MMM yyyy")
        Printer.CurrentX = Col2Left + 250
        Printer.CurrentY = YPos + Printer.TextHeight("X") + 200
        Printer.Print PrintMatter

        PrintMatter = Grid.TextMatrix(RowNo, 5)
        Printer.CurrentX = RightMargin - Printer.TextWidth(PrintMatter) - 1200
        Printer.CurrentY = YPos + Printer.TextHeight("X") + 200
        Printer.Print PrintMatter
        '''''''''''''''''''''''''''''''''''''

        PrintMatter = "Gross Deposit Amount"
        Printer.CurrentX = RightMargin - Printer.TextWidth(PrintMatter) - 1200
        Printer.CurrentY = YPos
        Printer.Print PrintMatter

        Printer.Line (Col1Left, YPos + Printer.TextHeight("X") + 450)-(RightMargin, YPos + Printer.TextHeight("X") + 450)
        YDum = YPos + Printer.TextHeight("X") + 450
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim DiffVal As Single
        DiffVal = Val(Val(Val(RightMargin) - Val(Col2Left)) / 3)

        YPos = YPos + Printer.TextHeight("X") + 500
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
        PrintMatter = Val(Grid.TextMatrix(RowNo, 4))
        Printer.CurrentY = YPos + Printer.TextHeight("X") + 100
        Printer.Print PrintMatter

        Printer.CurrentX = Col1Left + 100 'Printing Customer Code''''''''''''
        PrintMatter = Grid.TextMatrix(RowNo, 1)   ' ModInit.CCode
        Printer.CurrentY = YPos + Printer.TextHeight("X") + 100
        Printer.Print PrintMatter

        Printer.CurrentX = Col1Left + Val(Val(Col2Left - Col1Left) / 2) + 100 'Printing PickUpPoint''''''''''''
        PrintMatter = "HEE" 'Grid.TextMatrix(RowNo, 2)  'ModInit.PPoint
        Printer.CurrentY = YPos + Printer.TextHeight("X") + 100
        Printer.Print PrintMatter

        Printer.CurrentX = Col2Left + 100 'Printing Pickup Location''''''''''''
        PrintMatter = Grid.TextMatrix(RowNo, 9)   ' ModInit.PLoc
        Printer.CurrentY = YPos + Printer.TextHeight("X") + 100
        Printer.Print PrintMatter

        Printer.CurrentX = Col2Left + Val(Val(DiffVal) * 2) + 100 'Printing Customer's Ref''''''''''''
        PrintMatter = "" '  ModInit.Cref
        Printer.CurrentY = YPos + Printer.TextHeight("X") + 100
        Printer.Print PrintMatter

        '''''''''''''''''''''''''''''''''''''''''''
        Printer.CurrentX = Col2Left + Val(Val(DiffVal) * 2) + 100
        PrintMatter = "Customer's Ref"
        Printer.CurrentY = YPos
        Printer.Print PrintMatter
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        YPos = YPos + Printer.TextHeight("X") + 500

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
        
        If intCopies = 4 Then
            Printer.CurrentY = YPos
            PrintMatter = "Check No."
            Printer.CurrentX = Col1Left + 100
            Printer.Print PrintMatter
            
            Printer.CurrentY = YPos
            PrintMatter = "Barcode No"
            Printer.CurrentX = Col1Left + (Col3Left - Col1Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter
            
            Printer.CurrentY = YPos
            PrintMatter = "Check Date"
            Printer.CurrentX = Col3Left - (Printer.TextWidth(PrintMatter)) - 100
            Printer.Print PrintMatter
        Else
            Printer.CurrentY = YPos
            PrintMatter = "Check No."
            Printer.CurrentX = Col1Left + (Col2Left - Col1Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter
    
            Printer.CurrentY = YPos
            PrintMatter = "Check Date"
            Printer.CurrentX = Col2Left + (Col3Left - Col2Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter
        End If

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
                nxtrow = 300 * (Y - (PrintLines * (X - 1)))
            Else
                nxtrow = 300 * Y
            End If

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Val(Y)
            Printer.CurrentX = LeftMargin + (Col1Left - LeftMargin) / 2 - (Printer.TextWidth(PrintMatter) / 2)
            Printer.Print PrintMatter
            
            If intCopies = 4 Then
                Printer.CurrentY = YPos + nxtrow
                PrintMatter = CMSGrid.TextMatrix(Y, 1)
                Printer.CurrentX = Col1Left + 100 '(Col2Left - Col1Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
                Printer.Print PrintMatter
                
                Printer.CurrentY = YPos + nxtrow
                PrintMatter = CMSGrid.TextMatrix(Y, 18)
                Printer.CurrentX = Col1Left + (Col3Left - Col1Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
                Printer.Print PrintMatter
    
                Printer.CurrentY = YPos + nxtrow
                PrintMatter = CMSGrid.TextMatrix(Y, 2)
                Printer.CurrentX = Col3Left - (Printer.TextWidth(PrintMatter)) - 100
                Printer.Print PrintMatter
            Else
                Printer.CurrentY = YPos + nxtrow
                PrintMatter = CMSGrid.TextMatrix(Y, 1)
                Printer.CurrentX = Col1Left + (Col2Left - Col1Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
                Printer.Print PrintMatter
    
                Printer.CurrentY = YPos + nxtrow
                PrintMatter = CMSGrid.TextMatrix(Y, 2)
                Printer.CurrentX = Col2Left + (Col3Left - Col2Left) / 2 - (Printer.TextWidth(PrintMatter) / 2)
                Printer.Print PrintMatter
            End If

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = CMSGrid.TextMatrix(Y, 7)
            Printer.CurrentX = Col3Left + 100
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = Left(CMSGrid.TextMatrix(Y, 15), 45) 'Left(cmsgrid.TextMatrix(Y, 15), 40)
            Printer.CurrentX = Col4Left + 100
            Printer.Print PrintMatter

            Printer.CurrentY = YPos + nxtrow
            PrintMatter = CMSGrid.TextMatrix(Y, 3)
            Printer.CurrentX = RightMargin - (Printer.TextWidth(PrintMatter)) - 150
            Printer.Print PrintMatter

            If Y <> ENDNUM Then
                Printer.Line (LeftMargin, YPos + nxtrow + Printer.TextHeight("X") + 50)-(RightMargin, YPos + nxtrow + Printer.TextHeight("X") + 50)
            End If
        Next Y

        YPos = YPos + nxtrow + Printer.TextHeight("X") + 50
        Printer.Line (LeftMargin, YPos)-(RightMargin, YPos)

        Printer.Line (Col1Left, YStart1)-(Col1Left, YPos)
        If intCopies = 4 Then
        Else
            Printer.Line (Col2Left, YStart1)-(Col2Left, YPos)
        End If
        Printer.Line (Col3Left, YStart1)-(Col3Left, YPos)
        Printer.Line (Col4Left, YStart1)-(Col4Left, YPos)
        Printer.Line (Col5Left, YStart1)-(Col5Left, YPos)
'        Printer.Line (Col6Left, YStart1)-(Col6Left, YPos)
        Printer.Line (LeftMargin, YStart)-(LeftMargin, YPos)
        Printer.Line (RightMargin, YStart)-(RightMargin, YPos)

        YPos = YPos + 100

        If X = TPAGES Then

            Printer.CurrentY = YPos
            PrintMatter = "Please ensure all checks are drawn in favour of 'CITIBANK N.A. A/C COMPANY NAME' "
            Printer.CurrentX = LeftMargin + 100
            Printer.Print PrintMatter

            Printer.CurrentY = YPos
            PrintMatter = "TOTAL ->"
            Printer.CurrentX = Col5Left - Printer.TextWidth(PrintMatter) - 150
            Printer.Print PrintMatter

            Printer.CurrentY = YPos
            PrintMatter = Format(Grid.TextMatrix(RowNo, 5), "#,##,##,##0.00")
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

'        If OptLocal.Value = True Then
            If intCopies = 4 Then
                PrintMatter = "Customer Copy"
            ElseIf intCopies = 3 Then
                PrintMatter = "Co-Ordinator Copy"
            Else
                PrintMatter = "Citibank Copy"
            End If
'        Else
'            If intCopies = 3 Then
'                PrintMatter = "Customer Copy"
'            ElseIf intCopies = 2 Then
'                PrintMatter = "Co-Ordinator Copy"
'            Else
'                PrintMatter = "Citibank Copy"
'            End If
'        End If

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



Private Sub cmdPrint_Click()

End Sub



Private Sub cmdOutstation_Click()
On Error Resume Next
With FrmCMSDeposit_New_Sub1
    .FillCompany Format(dtSlip.Value, "ddMMyyyy"), 3
    
    If Val(Me.lblChqCnt_NonMICR.Caption) > 0 Then
        .mnuExportForBank.Enabled = False
        .cmdPrint.Enabled = False
    End If
    .Show vbModal
End With
End Sub

Private Sub cmdShowICICI_Click()
On Error Resume Next
With FrmCMSDeposit_New_Sub1
    .FillCompany Format(dtSlip.Value, "ddMMyyyy"), 1
    
    If Val(Me.lblChqCnt_NonMICR.Caption) > 0 Then
        .mnuExportForBank.Enabled = False
        .cmdPrint.Enabled = False
    End If
    .Show vbModal
End With
End Sub

Private Sub cmdShowNonMICR_Click()
On Error Resume Next
With FrmCMSDeposit_New_Sub1
    .FillCompany Format(dtSlip.Value, "ddMMyyyy"), 4
    
    .Show vbModal
End With
End Sub



Private Sub cmsShowNonICICI_Click()
On Error Resume Next
With FrmCMSDeposit_New_Sub1
    .FillCompany Format(dtSlip.Value, "ddMMyyyy"), 2
    
    If Val(Me.lblChqCnt_NonMICR.Caption) > 0 Then
        .mnuExportForBank.Enabled = False
        .cmdPrint.Enabled = False
    End If
    .Show vbModal
End With
End Sub

Private Sub Form_Load()
On Error Resume Next
dtSlip.Value = DateAdd("d", 1, Date)
Me.lblChqCnt_ICICI.Caption = "0"
Me.lblChqCnt_NonICICI.Caption = "0"
Me.lblChqCnt_NonMICR.Caption = "0"
Me.lblChqCnt_Outstation.Caption = "0"
End Sub

