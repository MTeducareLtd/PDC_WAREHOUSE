VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmCMSDeposit_Sub4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CMS Slip Generator"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   13125
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkGenerate 
      Caption         =   "Generate CMS"
      Height          =   195
      Left            =   4200
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdLockCMs 
      Caption         =   "&Lock CMS"
      Height          =   375
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Export &AS Copy"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print &AS Copy"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   120
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print CMS"
      Height          =   375
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   1695
   End
   Begin ComctlLib.ProgressBar pgBar 
      Height          =   255
      Left            =   8160
      TabIndex        =   7
      Top             =   240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
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
      Format          =   126025731
      CurrentDate     =   39310
   End
   Begin MSFlexGridLib.MSFlexGrid CMSGrid 
      Height          =   1050
      Left            =   240
      TabIndex        =   8
      Top             =   5880
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1852
      _Version        =   393216
      Cols            =   19
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
      FormatString    =   $"FrmCMSDeposit_Sub4.frx":0000
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
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5250
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   9260
      _Version        =   393216
      Cols            =   11
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
      FormatString    =   $"FrmCMSDeposit_Sub4.frx":0100
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
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   10800
      TabIndex        =   11
      Top             =   480
      Width           =   2055
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
      TabIndex        =   6
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "FrmCMSDeposit_Sub4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdGenerate_Click()
On Error GoTo ErrExit
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

'Find all centres
Dim SerStr As String
Dim RecCnt As Integer
Dim CMSStatus, CMSSlipNo, CMSSlipCode As String

Grid.Rows = 1

Dim ChqCnt As Integer
Dim ChqAmt As Double

If Me.dtSlip.Value > Date And Me.chkGenerate.Value = vbChecked Then
    SerStr = "Select distinct misinstitutecode as InstituteCode ,   liccode  from ASPDC_DispatchSlipDetails ADSD inner join ASPDC_DispatchSlip ADS on adsd.dispatchslipcode= ads.dispatchslipcode " & _
             "where CentreChequeDate<='" & Format(Me.dtSlip.Value, "dd Mmm yyyy") & "' and ccchqidno is not null and ((cmsdoneflag =0 or cmsdoneflag is null) or (cmsdoneflag =1 and cmsslipdate ='" & Format(dtSlip.Value, "dd Mmm yyyy") & "')) order by institutecode, liccode"
    rs1.Open SerStr, cn1, adOpenStatic, adLockReadOnly
    If Not (rs1.BOF And rs1.EOF) Then
        RecCnt = rs1.RecordCount
        Me.lblTot.Caption = "0.00"
        pgBar.Min = 0
        pgBar.Value = 0
        pgBar.Max = RecCnt
    
        Do While Not rs1.EOF
            CMSStatus = ""
            CMSSlipNo = ""
            'Add entry in ASPDC_CMS_CentreLog
    '        If rs1!InstituteCode = "600" And rs1!LicCode = "31" Then
    '            MsgBox ""
    '        End If
            
            rs2.Open "Select * from ASPDC_CMS_CentreLog where institutecode ='" & rs1!InstituteCode & "' and Liccode ='" & rs1!LicCode & "' and CMSDate ='" & Format(dtSlip.Value, "dd Mmm yyyy") & "' order by FinalChequeCnt desc", cn1, adOpenDynamic, adLockPessimistic
            If (rs2.BOF And rs2.EOF) Then
                If Me.dtSlip.Value > Date Then
                    rs2.AddNew
                    rs2!InstituteCode = rs1!InstituteCode
                    rs2!LicCode = rs1!LicCode
                    rs2!CMSDate = dtSlip.Value
                    rs2!CMSStatus = "Initialised"
                    rs2!FinalChequeCnt = 0
                    rs2!FinalChequeAmt = 0
                    
                    CMSStatus = "Initialised"
                    rs2!CMSSlipNo = ""
                    CMSSlipNo = ""
                    ChqCnt = 0
                    ChqAmt = 0
                    
                    rs2.Update
                Else
                    CMSSlipNo = ""
                    ChqCnt = 0
                    ChqAmt = 0
                End If
            Else
                CMSStatus = rs2!CMSStatus
                CMSSlipNo = rs2!CMSSlipNo
                CMSSlipCode = Val(rs2!CMSSlipNo)
                ChqCnt = Val("" & rs2!FinalChequeCnt)
                ChqAmt = Val("" & rs2!FinalChequeAmt)
            End If
            rs2.Close
            
            
            If (Me.dtSlip.Value <= Date And CMSStatus <> "Initialised") Or (Me.dtSlip.Value > Date) Then
                Grid.Rows = Grid.Rows + 1
                Grid.TextMatrix(Grid.Rows - 1, 6) = "" & rs1!InstituteCode
                Grid.TextMatrix(Grid.Rows - 1, 7) = "" & rs1!LicCode
                Grid.TextMatrix(Grid.Rows - 1, 8) = CMSStatus
                Grid.TextMatrix(Grid.Rows - 1, 3) = CMSSlipNo
                Grid.TextMatrix(Grid.Rows - 1, 0) = CMSSlipCode
                Grid.TextMatrix(Grid.Rows - 1, 4) = ChqCnt
                Grid.TextMatrix(Grid.Rows - 1, 5) = ChqAmt
                
                rs2.Open "Select * from c008_centers where Target_center_code ='" & rs1!InstituteCode & rs1!LicCode & "'", cn1, adOpenDynamic, adLockReadOnly
                If Not (rs2.BOF And rs2.EOF) Then
                    Grid.TextMatrix(Grid.Rows - 1, 1) = "" & rs2!CMS_Company_Code
                    Grid.TextMatrix(Grid.Rows - 1, 2) = "" & rs2!CMS_Pick_Up_Point
                    Grid.TextMatrix(Grid.Rows - 1, 9) = "" & rs2!CMS_Pickup_City
                End If
                rs2.Close
                
                lblTot.Caption = Val(lblTot.Caption) + ChqAmt
            End If
            
            'If ChqCnt = 0 Then Grid.RowHeight(Grid.Rows - 1) = 0
            
            pgBar.Value = pgBar.Value + 1
    
            rs1.MoveNext
        Loop
    End If
    rs1.Close
    lblTot.Caption = Format(lblTot.Caption, "0.00")
    
    'Start generating Slips for Initialised items
    Dim Cnt As Integer
    For Cnt = 1 To Grid.Rows - 1
        If Grid.TextMatrix(Cnt, 8) = "Initialised" Or Grid.TextMatrix(Cnt, 8) = "SlipNoLocked" Then
            'Find cheques for these centre
            AddSlip Grid.TextMatrix(Cnt, 6) & Grid.TextMatrix(Cnt, 7), Cnt
            
            DoEvents
        End If
        
        If Grid.TextMatrix(Cnt, 4) = 0 Then Grid.RowHeight(Cnt) = 0
    Next
    
Else
    SerStr = "Select distinct InstituteCode, LicCode, CMSSlipNo from ASPDC_CMS_CentreLog where CMSDate ='" & Format(dtSlip.Value, "dd Mmm yyyy") & "' and isnull(FinalChequeCnt,0) >0 order by institutecode, liccode"

    rs1.Open SerStr, cn1, adOpenStatic, adLockReadOnly
    If Not (rs1.BOF And rs1.EOF) Then
        RecCnt = rs1.RecordCount
        Me.lblTot.Caption = "0.00"
        pgBar.Min = 0
        pgBar.Value = 0
        pgBar.Max = RecCnt
    
        Do While Not rs1.EOF
            CMSStatus = ""
            CMSSlipNo = ""
            'Add entry in ASPDC_CMS_CentreLog
            
            rs2.Open "Select * from ASPDC_CMS_CentreLog where institutecode ='" & rs1!InstituteCode & "' and Liccode ='" & rs1!LicCode & "' and CMSSlipNo ='" & rs1!CMSSlipNo & "' order by FinalChequeCnt desc", cn1, adOpenDynamic, adLockPessimistic
            If Not (rs2.BOF And rs2.EOF) Then
                CMSStatus = rs2!CMSStatus
                CMSSlipNo = rs2!CMSSlipNo
                CMSSlipCode = Val(rs2!CMSSlipNo)
                ChqCnt = Val("" & rs2!FinalChequeCnt)
                ChqAmt = Val("" & rs2!FinalChequeAmt)
            End If
            rs2.Close
            
            
            If (Me.dtSlip.Value <= Date And CMSStatus <> "Initialised") Or (Me.dtSlip.Value > Date) Then
                Grid.Rows = Grid.Rows + 1
                Grid.TextMatrix(Grid.Rows - 1, 6) = "" & rs1!InstituteCode
                Grid.TextMatrix(Grid.Rows - 1, 7) = "" & rs1!LicCode
                Grid.TextMatrix(Grid.Rows - 1, 8) = CMSStatus
                Grid.TextMatrix(Grid.Rows - 1, 3) = CMSSlipNo
                Grid.TextMatrix(Grid.Rows - 1, 0) = CMSSlipCode
                Grid.TextMatrix(Grid.Rows - 1, 4) = ChqCnt
                Grid.TextMatrix(Grid.Rows - 1, 5) = ChqAmt
                
                rs2.Open "Select * from c008_centers where Target_center_code ='" & rs1!InstituteCode & rs1!LicCode & "'", cn1, adOpenDynamic, adLockReadOnly
                If Not (rs2.BOF And rs2.EOF) Then
                    Grid.TextMatrix(Grid.Rows - 1, 1) = "" & rs2!CMS_Company_Code
                    Grid.TextMatrix(Grid.Rows - 1, 2) = "" & rs2!CMS_Pick_Up_Point
                    Grid.TextMatrix(Grid.Rows - 1, 9) = "" & rs2!CMS_Pickup_City
                End If
                rs2.Close
                
                lblTot.Caption = Val(lblTot.Caption) + ChqAmt
            End If
            
            'If ChqCnt = 0 Then Grid.RowHeight(Grid.Rows - 1) = 0
            
            pgBar.Value = pgBar.Value + 1
    
            rs1.MoveNext
        Loop
    End If
    rs1.Close
    lblTot.Caption = Format(lblTot.Caption, "0.00")
    
    
End If
'---------------

cn1.Close

MsgBox RecCnt & " CMS Slips generated.", vbInformation + vbOKOnly

cmdGenerate.Enabled = False
Exit Sub
ErrExit:
MsgBox "Error: " & Err.Description, vbCritical + vbOKOnly

End Sub


Private Sub AddSlip(InstituteCode As String, RowNo As Integer)
On Error Resume Next
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

'Generate SlipId
Dim SlipId, DepositSlipNoPre As String

If Grid.TextMatrix(RowNo, 8) = "Initialised" Then
    If dtSlip.Value <= DateValue("31 Mar " & Year(Date)) Then
        DepositSlipNoPre = InstituteCode & Val(Right(Year(Date), 1))
    Else
        DepositSlipNoPre = InstituteCode & (Val(Right(Year(Date), 1)) + 1)
    End If
    
    If Trim(DepositSlipNoPre) <> "" Then
        rs1.Open "Select max(right(SlipNo,4)) As MainVal from DepositSlip where Left(SlipNo,6)='" & DepositSlipNoPre & "'  and len(slipno) = 10 and slipdate >='1 Jan 2013'", cn1, adOpenDynamic, adLockReadOnly
        If Not (rs1.BOF And rs1.EOF) Then
            DepositSlipNoPost = Val("" & rs1.Fields("MainVal").Value)
        End If
        rs1.Close
        
        If Val(DepositSlipNoPost) = 0 Then DepositSlipNoPost = "3000"
        
        If Val(DepositSlipNoPost) < 3000 Then DepositSlipNoPost = Val(DepositSlipNoPost) + 3000
    
        SlipId = DepositSlipNoPre & Format(Val(DepositSlipNoPost) + 1, "0000")
    Else
        SlipId = ""
    End If
    Grid.TextMatrix(RowNo, 3) = SlipId
    Grid.TextMatrix(RowNo, 0) = SlipId
Else
    SlipId = Grid.TextMatrix(RowNo, 3)
End If

DoEvents

'Generate SlipNo
Dim SlipNo As String
SlipNo = SlipId

If SlipNo = "" Then
    cn1.Close
    Exit Sub
End If

Dim ChqCnt As Long
ChqCnt = 0

Dim TotalChqVal As Double
TotalChqVal = 0

rs1.Open "Select * from DepositSlip where SlipCode =" & SlipId & " and CentreCode ='" & InstituteCode & "'", cn1, adOpenDynamic, adLockOptimistic
If (rs1.BOF And rs1.EOF) Then
    rs1.AddNew
    rs1.Fields("SlipCode").Value = Val(SlipId)
    rs1.Fields("SlipNo").Value = SlipNo
    rs1.Fields("SlipDate").Value = Me.dtSlip.Value
    rs1.Fields("ChqCnt").Value = ChqCnt     'Default value is 0
    rs1.Fields("ChqAmt").Value = TotalChqVal    'Default value is 0
    rs1.Fields("PrintFlag").Value = 0
    rs1.Fields("CentreCode").Value = InstituteCode
    rs1.Fields("DepositSlipType").Value = "Bank"
    rs1.Fields("MISInstituteCode").Value = Left(InstituteCode, 3)
    rs1.Update
End If
rs1.Close

DoEvents

If Grid.TextMatrix(RowNo, 8) = "Initialised" Then
    rs1.Open "Select * from ASPDC_CMS_CentreLog where institutecode ='" & Left(InstituteCode, 3) & "' and Liccode ='" & Right(InstituteCode, 2) & "' and CMSDate ='" & Format(dtSlip.Value, "dd Mmm yyyy") & "'", cn1, adOpenDynamic, adLockPessimistic
    If Not (rs1.BOF And rs1.EOF) Then
        rs1!CMSSlipNo = SlipNo
        rs1!CMSStatus = "SlipNoLocked"
        Grid.TextMatrix(RowNo, 8) = "SlipNoLocked"
        rs1.Update
    End If
    rs1.Close
End If

DoEvents

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

Dim rs3 As ADODB.Recordset
Set rs3 = New ADODB.Recordset

Dim rs4 As ADODB.Recordset
Set rs4 = New ADODB.Recordset

Dim CCnt As Integer
CCnt = 0

Dim ChqIDNo As String
Dim ChqAmt As Double

'save deposit slip details
If Grid.TextMatrix(RowNo, 8) = "SlipNoLocked" Then

    rs1.Open "Delete from DepositSlipDetails where SlipCode =" & SlipId & " and CentreCode ='" & InstituteCode & "'", cn1, adOpenDynamic, adLockPessimistic
    
    SerStr = "Select CCChequeNo, CentreChequeDate, CCChequeAmt, ChqIDNo, isnull(HoldFlag,0) as HoldFlag, isnull(ReturnFlag,0) as ReturnFlag , isnull(CMSDoneFlag,0) as CMSDoneFlag from ASPDC_DispatchSlipDetails ADSD inner join ASPDC_DispatchSlip ADS on adsd.dispatchslipcode= ads.dispatchslipcode " & _
             "where CentreChequeDate <='" & Format(dtSlip.Value, "dd Mmm yyyy") & "' and CCChequeDate <='" & Format(dtSlip.Value, "dd Mmm yyyy") & "' and (cmsdoneflag =0 or cmsdoneflag is null) and ccchqidno is not null and misinstitutecode ='" & Left(InstituteCode, 3) & "' and liccode ='" & Right(InstituteCode, 2) & "' and (HoldFlag =0 or HoldFlag is Null) and (ReturnFlag =0 or ReturnFlag is Null)"
    
    rs2.Open SerStr, cn1, adOpenDynamic, adLockReadOnly
    If Not (rs2.BOF And rs2.EOF) Then
        
        
        rs1.Open "Select * from DepositSlipDetails where SlipCode =" & SlipId & " and CentreCode ='" & InstituteCode & "'", cn1, adOpenDynamic, adLockPessimistic
        rs2.MoveFirst
        Do While Not rs2.EOF
            If rs2!HoldFlag = 0 And rs2!ReturnFlag = 0 And rs2!CMSDoneFlag = 0 Then
                'Check if chqidno exists in studentpayment
                rs3.Open "select sp.yearname, sp.rcptcode, sp.batchcode, sp.rollno, sp.SBEntryCode from studentpayment sp inner join StudentBatch sb on sp.yearname = sb.yearname and sp.institutecode = sb.institutecode and sp.liccode = sb.liccode and sp.sbentrycode = sb.sbentrycode where sp.chequeidno ='" & rs2!ChqIDNo & "' and sp.recorddelflag =0 and sb.status =1", cn1, adOpenDynamic, adLockReadOnly
                
                If Not (rs3.BOF And rs3.EOF) Then
                    'Check if sbentrycode is available for the student
                    'rs3.Open "select sp.yearname, sp.rcptcode, sp.batchcode, sp.rollno, sp.SBEntryCode from studentpayment sp inner join studentbatch sb on sp.yearname = sb.yearname and sp.institutecode = sb.institutecode and sp.sbentrycode = sb.sbentrycode where sp.chequeidno ='" & rs2!ChqIDNo & "' and sb.status =1 and sb.pendingflag =0 and sp.recorddelflag =0", cn1, adOpenDynamic, adLockReadOnly
                    
                    CCnt = CCnt + 1
                    rs1.AddNew
                    rs1.Fields("SlipCode").Value = Val(SlipId)
                    rs1.Fields("EntryCode").Value = CCnt
                    rs1.Fields("YearName").Value = rs3!YearName
                    rs1.Fields("BatchCode").Value = rs3!BatchCode
                    rs1.Fields("RollNo").Value = rs3!RollNo
                    rs1.Fields("ChequeNo").Value = rs2!CCChequeNo
            
                    rs1.Fields("ChequeDate").Value = rs2!CentreChequeDate
                    rs1.Fields("ChequeStatus").Value = "Deposited"
                    rs1.Fields("ChequeAmt").Value = rs2!CCChequeAmt
                    rs1.Fields("RcptCode").Value = rs3!RcptCode
                    rs1.Fields("SBEntryCode").Value = rs3!SBEntryCode
                    rs1.Fields("CentreCode").Value = InstituteCode
                    rs1.Fields("ChequeIdNo").Value = rs2!ChqIDNo
                    rs1.Fields("DownloadFlag").Value = "True"
                    
                    ChqAmt = ChqAmt + Val(rs2!CCChequeAmt)
                    rs1.Update
                    
                    rs4.Open "Update ASPDC_DispatchSlipDetails set cmsdoneflag =1, CMSSlipNo ='" & SlipNo & "', CMSSlipCode =" & SlipId & ", CMSSlipDate ='" & Format(dtSlip.Value, "dd Mmm yyyy") & "' where ChqIDNo ='" & rs2!ChqIDNo & "' and CCChequeNo ='" & rs2!CCChequeNo & "'", cn1, adOpenDynamic, adLockPessimistic
                    
                End If
                rs3.Close
            End If
            
            DoEvents
            
            rs2.MoveNext
        Loop
        rs1.Close
    End If
    rs2.Close
    
    'Change flag to SlipGenerated
    rs1.Open "Select * from ASPDC_CMS_CentreLog where institutecode ='" & Left(InstituteCode, 3) & "' and Liccode ='" & Right(InstituteCode, 2) & "' and CMSDate ='" & Format(dtSlip.Value, "dd Mmm yyyy") & "'", cn1, adOpenDynamic, adLockPessimistic
    If Not (rs1.BOF And rs1.EOF) Then
        rs1!FinalChequeCnt = CCnt
        rs1!FinalChequeAmt = ChqAmt
        rs1!CMSStatus = "SlipGenerated"
        Grid.TextMatrix(RowNo, 8) = "SlipGenerated"
        Grid.TextMatrix(RowNo, 4) = CCnt
        Grid.TextMatrix(RowNo, 5) = ChqAmt
        rs1.Update
    End If
    rs1.Close
    
    rs1.Open "Update DepositSlip set ChqCnt =" & CCnt & ", ChqAmt =" & ChqAmt & ", PrintFlag =1 where SlipCode =" & SlipId & " and CentreCode ='" & InstituteCode & "'", cn1, adOpenDynamic, adLockOptimistic
   
End If

cn1.Close
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

Private Sub cmdLockCMs_Click()
On Error Resume Next
Dim Cnt As Integer

For Cnt = 1 To Grid.Rows - 1
    If Grid.TextMatrix(Cnt, 8) = "SlipGenerated" And Grid.TextMatrix(Cnt, 10) = "Print" Then
        VerifyGridCMS Cnt
        
       
        Grid.TextMatrix(Cnt, 10) = "Verified"
        Grid.RowHeight(Cnt) = 0
    End If
Next
End Sub


Private Sub VerifyGridCMS(RowNo As Integer)
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

Err.Clear
rs1.Open "Delete from DepositSlipDetails where SlipCode =" & Grid.TextMatrix(RowNo, 3), cn1, adOpenDynamic, adLockPessimistic

SerStr = "Select CCChequeNo, CentreChequeDate, CCChequeAmt, ChqIDNo from ASPDC_DispatchSlipDetails ADSD inner join ASPDC_DispatchSlip ADS on adsd.dispatchslipcode= ads.dispatchslipcode " & _
         "where cmsslipno ='" & Grid.TextMatrix(RowNo, 3) & "'"
   
rs2.Open SerStr, cn1, adOpenDynamic, adLockReadOnly
If Not (rs2.BOF And rs2.EOF) Then
    rs1.Open "Select * from DepositSlipDetails where SlipCode =" & Grid.TextMatrix(RowNo, 3) & "", cn1, adOpenDynamic, adLockPessimistic
    rs2.MoveFirst
    CCnt = 1
    Do While Not rs2.EOF
        'Check if chqidno exists in studentpayment
        rs3.Open "select sp.yearname, sp.rcptcode, sp.batchcode, sp.rollno, sp.SBEntryCode from studentpayment sp inner join studentbatch sb on sp.yearname = sb.yearname and sp.institutecode = sb.institutecode and sp.sbentrycode = sb.sbentrycode where sp.chequeidno ='" & rs2!ChqIDNo & "' and sb.status =1 and sb.pendingflag =0 and sp.recorddelflag =0", cn1, adOpenDynamic, adLockReadOnly
        If Not (rs3.BOF And rs3.EOF) Then
            rs1.AddNew
            rs1.Fields("SlipCode").Value = Val(Grid.TextMatrix(RowNo, 3))
            rs1.Fields("EntryCode").Value = CCnt
            rs1.Fields("YearName").Value = rs3!YearName
            rs1.Fields("BatchCode").Value = rs3!BatchCode
            rs1.Fields("RollNo").Value = rs3!RollNo
            rs1.Fields("ChequeNo").Value = rs2!CCChequeNo
    
            rs1.Fields("ChequeDate").Value = rs2!CentreChequeDate
            rs1.Fields("ChequeStatus").Value = "Deposited"
            rs1.Fields("ChequeAmt").Value = rs2!CCChequeAmt
            rs1.Fields("RcptCode").Value = rs3!RcptCode
            rs1.Fields("SBEntryCode").Value = rs3!SBEntryCode
            rs1.Fields("CentreCode").Value = Left(Grid.TextMatrix(RowNo, 3), 5)
            rs1.Fields("ChequeIdNo").Value = rs2!ChqIDNo
            rs1.Fields("DownloadFlag").Value = "True"
            
            rs1.Update
            CCnt = CCnt + 1
        End If
        rs3.Close
        
        
        DoEvents
        
        Err.Clear
        rs2.MoveNext
        If Err.Number > 0 Then GoTo ErrExit
    Loop
End If
rs2.Close

ErrExit:
cn1.Close
End Sub


Private Sub cmdPrint_Click()
On Error GoTo ErrHandler
cmd.CancelError = True

On Error GoTo ErrHandler
Dim NumCopies As Integer

'Set flags
cmd.Flags = &H100000 Or &H4

' Display the Print dialog box.
cmd.ShowPrinter
NumCopies = cmd.Copies
Printer.Orientation = cdlLandscape

Dim Cnt As Integer

For Cnt = 1 To Grid.Rows - 1
    If Grid.TextMatrix(Cnt, 8) = "SlipGenerated" And Grid.TextMatrix(Cnt, 10) = "Print" Then
        FillGridCMS Cnt
        
        'Check if companycode exists for this slip
        If Grid.TextMatrix(Cnt, 1) = "" Then
            Grid.TextMatrix(Cnt, 10) = "CCode Blank"
        Else
            'Check if cheque cnt in cms grid and grid are matching
            If CMSGrid.Rows - 1 = Val(Grid.TextMatrix(Cnt, 4)) Then
                PrintRes Cnt, NumCopies
                Grid.TextMatrix(Cnt, 10) = "Printed"
                Grid.RowHeight(Cnt) = 0
            Else
                MsgBox "Mismatch found in Cheque Count.", vbCritical + vbOKOnly
                Grid.TextMatrix(Cnt, 10) = "Mismatch"
            End If
        End If
        
    End If
Next

MsgBox "CMS Printing completed successfully.", vbInformation + vbOKOnly

ErrHandler:
End Sub

Private Sub FillGridCMS(RowNo As Integer)
On Error Resume Next
CMSGrid.Rows = 1

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

Err.Clear
rs1.Open "Select * from DepositSlipDetails where SlipCode =" & Grid.TextMatrix(RowNo, 3) & "", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        CMSGrid.Rows = CMSGrid.Rows + 1
        CMSGrid.TextMatrix(CMSGrid.Rows - 1, 0) = "Deposited"
        CMSGrid.TextMatrix(CMSGrid.Rows - 1, 1) = rs1!ChequeNo
        CMSGrid.TextMatrix(CMSGrid.Rows - 1, 2) = Format(rs1.Fields("ChequeDate").Value, "ddMmyyyy")
        CMSGrid.TextMatrix(CMSGrid.Rows - 1, 3) = Format(rs1.Fields("ChequeAmt").Value, "0.00")
        CMSGrid.TextMatrix(CMSGrid.Rows - 1, 4) = "002"
        CMSGrid.TextMatrix(CMSGrid.Rows - 1, 5) = "002"
        CMSGrid.TextMatrix(CMSGrid.Rows - 1, 6) = "10"

        CMSGrid.TextMatrix(CMSGrid.Rows - 1, 10) = rs1!YearName
        CMSGrid.TextMatrix(CMSGrid.Rows - 1, 11) = rs1!BatchCode
        CMSGrid.TextMatrix(CMSGrid.Rows - 1, 12) = rs1!RollNo
        CMSGrid.TextMatrix(CMSGrid.Rows - 1, 13) = rs1!RcptCode

        CMSGrid.TextMatrix(CMSGrid.Rows - 1, 14) = Format(rs1!ChequeDate, "dd MMM yyyy")


        rs3.Open "Select Title, FirstName, MidName, LastName, Sex, SBEntryCode from Student inner join StudentBatch on Student.Yearname = StudentBatch.YearName and Student.InstituteCode = StudentBatch.InstituteCode and Student.LicCode = StudentBatch.LicCode and Student.StudentCode = StudentBatch.StudentCode where StudentBatch.RecordDelFlag = 0 and StudentBatch.BatchCode ='" & rs1.Fields("BatchCode").Value & "' and StudentBatch.RollNo =" & rs1.Fields("RollNo").Value, cn1, adOpenDynamic, adLockReadOnly
        If Not (rs3.BOF And rs3.EOF) Then
            CMSGrid.TextMatrix(CMSGrid.Rows - 1, 7) = rs3.Fields("FirstName").Value & " " & Left(rs3.Fields("MidName").Value, 1) & " " & rs3.Fields("LastName").Value
            CMSGrid.TextMatrix(CMSGrid.Rows - 1, 9) = rs3!Sex
        Else
            rs3.Close
            rs3.Open "Select Title, First_Name, Mid_Name, Last_Name, Stream_Desc from T000_Student_Personal_Data where cur_sb_code ='" & rs1.Fields("SBEntryCode").Value & "'", cn1, adOpenDynamic, adLockReadOnly
            If Not (rs3.BOF And rs3.EOF) Then
                CMSGrid.TextMatrix(CMSGrid.Rows - 1, 7) = rs3.Fields("First_Name").Value & " " & Left(rs3.Fields("Mid_Name").Value, 1) & " " & rs3.Fields("Last_Name").Value
                CMSGrid.TextMatrix(CMSGrid.Rows - 1, 9) = ""
                CMSGrid.TextMatrix(CMSGrid.Rows - 1, 8) = rs3!Stream_Desc
            End If
        End If
        rs3.Close
        CMSGrid.TextMatrix(CMSGrid.Rows - 1, 16) = rs1!SBEntryCode
        
        rs3.Open "Select BankName from StudentPayment where SBEntryCode ='" & rs1!SBEntryCode & "' and RcptCode ='" & rs1!RcptCode & "'", cn1, adOpenDynamic, adLockReadOnly
        If Not (rs3.BOF And rs3.EOF) Then
            CMSGrid.TextMatrix(CMSGrid.Rows - 1, 15) = "" & rs3!BankName
        End If
        rs3.Close

        rs3.Open "SELECT Streams.StreamName FROM StudentBatch INNER JOIN (Batches INNER JOIN Streams ON Batches.StreamCode = Streams.StreamCode) ON StudentBatch.BatchCode = Batches.BatchCode WHERE  StudentBatch.RecordDelFlag = 0 and StudentBatch.BatchCode='" & rs1.Fields("BatchCode").Value & "' AND StudentBatch.RollNo=" & rs1.Fields("RollNo").Value, cn1, adOpenDynamic, adLockReadOnly
        If Not (rs3.BOF And rs3.EOF) Then
            CMSGrid.TextMatrix(CMSGrid.Rows - 1, 8) = rs3!StreamName
        End If
        rs3.Close
        
        CMSGrid.TextMatrix(CMSGrid.Rows - 1, 17) = rs1!ChequeIdNo
        
        rs3.Open "Select CCCHQIdNo from ASPDC_DispatchSlipDetails where CHQIdNo ='" & rs1!ChequeIdNo & "'", cn1, adOpenDynamic, adLockReadOnly
        If Not (rs3.BOF And rs3.EOF) Then
            CMSGrid.TextMatrix(CMSGrid.Rows - 1, 18) = rs3!CCCHQIdNo
        End If
        rs3.Close
        
        Err.Clear
        rs1.MoveNext
        If Err.Number > 0 Then GoTo ErrExit
    Loop
End If
rs1.Close

CMSGrid.Col = 18
CMSGrid.ColSel = 18
CMSGrid.Sort = 1

ErrExit:
cn1.Close
End Sub


Private Sub Form_Load()
On Error Resume Next
dtSlip.Value = DateAdd("d", 1, Date)
Grid.Rows = 1
With Grid
    For Cnt = 0 To .Cols - 1
        .ColWidth(Cnt) = (.Width - 350) / 7
    Next
    .ColWidth(0) = 0
    .ColWidth(6) = 0
    .ColWidth(7) = 0
    .ColWidth(9) = 0
End With

End Sub

Private Sub Grid_DblClick()
On Error Resume Next
If Grid.Rows <= 1 Then Exit Sub

If Grid.TextMatrix(Grid.RowSel, 10) = "" Then
    Grid.TextMatrix(Grid.RowSel, 10) = "Print"
Else
    Grid.TextMatrix(Grid.RowSel, 10) = ""
End If
End Sub

Private Sub Label1_Click()

End Sub
