VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.MDIForm FrmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Acountech PDC Factory Manager"
   ClientHeight    =   7185
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11430
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pictDashboard 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   11370
      TabIndex        =   5
      Top             =   0
      Width           =   11430
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   0
         Top             =   0
      End
      Begin SHDocVwCtl.WebBrowser wbSend 
         Height          =   1335
         Left            =   19080
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
         ExtentX         =   2990
         ExtentY         =   2355
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
      Begin VB.Line Line6 
         BorderColor     =   &H00C0C0C0&
         X1              =   15000
         X2              =   15000
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   16920
         TabIndex        =   26
         Top             =   345
         Width           =   1320
      End
      Begin VB.Label lblBounceCheque 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   16920
         TabIndex        =   25
         Top             =   120
         Width           =   1320
      End
      Begin VB.Label lblRemovedChq 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   13440
         TabIndex        =   24
         Top             =   560
         Width           =   1320
      End
      Begin VB.Label lblCorrectedChq 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   13440
         TabIndex        =   23
         Top             =   340
         Width           =   1320
      End
      Begin VB.Label lblDispatchSlip 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   13440
         TabIndex        =   22
         Top             =   120
         Width           =   1320
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Deposited Cheques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   15240
         TabIndex        =   21
         Top             =   340
         Width           =   1665
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Bounced Cheques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   15240
         TabIndex        =   20
         Top             =   120
         Width           =   1560
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Removed Cheques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11760
         TabIndex        =   19
         Top             =   555
         Width           =   1605
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Corrected Cheques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11760
         TabIndex        =   18
         Top             =   345
         Width           =   1635
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Dispatch Slip"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11760
         TabIndex        =   17
         Top             =   120
         Width           =   1140
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         X1              =   11520
         X2              =   11520
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         X1              =   8880
         X2              =   8880
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Bounce Mail Pending"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9120
         TabIndex        =   15
         Top             =   120
         Width           =   2205
      End
      Begin VB.Label lblPending_BouncedMail 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   9120
         TabIndex        =   14
         Top             =   360
         Width           =   2205
      End
      Begin VB.Label lblPending_Auth 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   6360
         TabIndex        =   13
         Top             =   360
         Width           =   2280
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Authorisation Pending"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6360
         TabIndex        =   12
         Top             =   120
         Width           =   2280
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   6120
         X2              =   6120
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   4200
         X2              =   4200
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   2160
         X2              =   2160
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Scanning Errors"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2400
         TabIndex        =   11
         Top             =   120
         Width           =   1665
      End
      Begin VB.Label lblError_Scanning 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   2400
         TabIndex        =   10
         Top             =   360
         Width           =   1665
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Entry Pending"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4440
         TabIndex        =   9
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblPending_Entry 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   4440
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblPending_Scan 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1890
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Scanning Pending"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1890
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   11370
      TabIndex        =   0
      Top             =   6840
      Width           =   11430
      Begin VB.Label lblLocation 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   4560
         TabIndex        =   4
         Top             =   60
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location "
         Height          =   195
         Left            =   3600
         TabIndex        =   3
         Top             =   60
         Width           =   660
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1080
         TabIndex        =   2
         Top             =   60
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   60
         Width           =   795
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "Inward"
      Begin VB.Menu mnuInwardEntry 
         Caption         =   "Inward Entry"
      End
      Begin VB.Menu mnuDispatchSlipEntry 
         Caption         =   "Dispatch Slip Entry "
      End
      Begin VB.Menu mnuDispatchSlipChqEntry 
         Caption         =   "Dispatch Slip Cheque Entry"
      End
      Begin VB.Menu mnuDispatchMICRScan 
         Caption         =   "Cheque - MICR Entry"
      End
      Begin VB.Menu mnuDeleteDispatchSlipEntry 
         Caption         =   "Delete Dispatch Slip Entry"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDispatchSlipAuthorisation 
         Caption         =   "Dispatch Slip Authorisation"
      End
      Begin VB.Menu mnuDispatchSlipCancellation 
         Caption         =   "Dispatch Slip Cancellation"
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBounceChequeEntry 
         Caption         =   "Bounce Cheque Inward Entry"
      End
   End
   Begin VB.Menu mnuOutward 
      Caption         =   "Outward"
      Begin VB.Menu mnuMakeCMS 
         Caption         =   "Make CMS"
      End
      Begin VB.Menu mnuReturnChequeEntry 
         Caption         =   "Return Cheque Entry"
      End
      Begin VB.Menu mnuBounceChequeReturnEntry 
         Caption         =   "Bounce Cheque Return Entry"
      End
      Begin VB.Menu mnuHoldChequeEntry 
         Caption         =   "Hold Cheque Entry"
      End
      Begin VB.Menu mnuSearchCheque 
         Caption         =   "Search Cheque"
      End
   End
   Begin VB.Menu mnuUtilities 
      Caption         =   "Utilities"
      Begin VB.Menu mnuSearchChequeNo 
         Caption         =   "Search Cheque on No"
      End
      Begin VB.Menu mnuActivatePDC 
         Caption         =   "Activate PDC"
      End
      Begin VB.Menu mnuDataProcessing 
         Caption         =   "Data Processing"
      End
      Begin VB.Menu mnuCMSSettings 
         Caption         =   "CMS Settings"
      End
      Begin VB.Menu mnuAutoMailBouncedCheque 
         Caption         =   "Auto Mail - Bounced Cheque"
      End
      Begin VB.Menu mnuMonthlyStockCheck 
         Caption         =   "Monthly Stock Checking"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuRepPendingSlipCollection 
         Caption         =   "Pending Slip Collection"
      End
      Begin VB.Menu mnuRepCMS 
         Caption         =   "CMS Report"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuReturnChequeRequests 
         Caption         =   "Return Cheque Requests"
      End
   End
   Begin VB.Menu mnuLogout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WaitFlag As Boolean


Private Sub MDIForm_Load()
ModInit.Init

If ModInit.PDCUserType = 1 Or ModInit.PDCUserType = 3 Or ModInit.PDCUserType = 2 Then  'Manager & Slip Entry
    mnuDispatchSlipEntry.Visible = True
Else
    mnuDispatchSlipEntry.Visible = False
End If

If ModInit.PDCUserType = 1 Or ModInit.PDCUserType = 2 Or ModInit.PDCUserType = 3 Then  'Manger & Cheque entry
    Me.mnuDispatchSlipChqEntry.Visible = True
Else
    Me.mnuDispatchSlipChqEntry.Visible = False
End If

If ModInit.PDCUserType = 1 Then 'Manager
    Me.mnuDispatchSlipAuthorisation.Visible = True
    Me.mnuDispatchSlipCancellation.Visible = True
    Me.mnuActivatePDC.Visible = True
Else
    Me.mnuDispatchSlipAuthorisation.Visible = False
    mnuDispatchSlipCancellation.Visible = False
    Me.mnuActivatePDC.Visible = False
End If

If ModInit.PDCUserType = 1 Or ModInit.PDCUserType = 4 Then  'Manager & CMS
    mnuMakeCMS.Visible = True
    mnuBounceChequeEntry.Visible = True
Else
    mnuMakeCMS.Visible = False
    mnuBounceChequeEntry.Visible = False
End If

If ModInit.LocationCode <> "001" Then
    mnuUtilities.Visible = False
    mnuDeleteDispatchSlipEntry.Visible = False
    mnuReports.Visible = False
    mnuReturnChequeEntry.Visible = False
    mnuHoldChequeEntry.Visible = False
    pictDashboard.Visible = False
Else
    If ModInit.PDCUserType = 1 Then
        mnuDeleteDispatchSlipEntry.Visible = True
    Else
        mnuDeleteDispatchSlipEntry.Visible = False
    End If
    pictDashboard.Visible = True
End If

Me.lblLocation.Caption = ModInit.LocationName
Me.lblUserName.Caption = ModInit.PDCUserName

UpdateDashboard
End Sub

Private Sub mnuActivatePDC_Click()
On Error Resume Next
FrmActivatePDC.Show vbModal
End Sub

Private Sub mnuAutoMailBouncedCheque_Click()
On Error Resume Next
With FrmAutoMailBouncedCheque
    .Show
    .WindowState = vbMaximized
    .ZOrder (0)
End With
End Sub

Private Sub mnuBounceChequeEntry_Click()
On Error Resume Next
With FrmBounceChequeEntry
    .Show
    .WindowState = vbMaximized
    .ZOrder (0)
End With
End Sub

Private Sub mnuBounceChequeReturnEntry_Click()
On Error Resume Next
With FrmReturnBounceChequeEntry
    .Show
    .WindowState = vbMaximized
    .ZOrder (0)
End With
End Sub

Private Sub mnuCMSSettings_Click()
On Error Resume Next
FrmCMSSettings.Show vbModal
End Sub

Private Sub mnuDataProcessing_Click()
On Error Resume Next
'ProcessData
Dim CMSSlipNo As String
CMSSlipNo = InputBox("Enter CMS Slip Number")

If Len(CMSSlipNo) > 0 And Len(CMSSlipNo) <> 10 Then
    MsgBox "Invalid CMS Slip Number.", vbCritical + vbOKOnly
Else
    RecreateDepositSlip CMSSlipNo, Left(CMSSlipNo, 5)
End If

End Sub

Private Sub ProcessData()
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

Dim SerStr As String
SerStr = "select  asd.CMSSlipNo, count(*) ChqInASD, ascm.FinalChequeCnt from ASPDC_DispatchSlipDetails asd inner join ASPDC_CMS_CentreLog ascm  on asd.cmsslipno = ascm.CMSSlipNo group by asd.CMSSlipNo , ascm.FinalChequeCnt having count(*)  > ascm.FinalChequeCnt"

rs1.Open SerStr, cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        'check if cheque count in Depositslipdetails is same as FinalChequeCnt
        rs2.Open "select count(*) as DSDChqCnt from DepositSlipDetails where SlipCode =" & Val(rs1!CMSSlipNo), cn1, adOpenDynamic, adLockReadOnly
        If Not (rs2.BOF And rs2.EOF) Then
            If rs2!DSDChqCnt = rs1!FinalChequeCnt Then
                'Remove cms flag from ASPDC_DispatchSlipDetails where cheques are not in Depositslipdetails
                rs3.Open "Select * from ASPDC_DispatchSlipDetails where CMSSlipNo ='" & rs1!CMSSlipNo & "'", cn1, adOpenDynamic, adLockPessimistic
                If Not (rs3.BOF And rs3.EOF) Then
                    Do While Not rs3.EOF
                        rs4.Open "Select * from DepositSlipDetails where SlipCode =" & Val(rs1!CMSSlipNo) & " and ChequeIdNo ='" & rs3!ChqIdNo & "'", cn1, adOpenDynamic, adLockReadOnly
                        If (rs4.BOF And rs4.EOF) Then
                            'Cheque not found
                            rs3!CMSDoneFlag = 0
                            rs3!CMSSlipCode = ""
                            rs3!CMSSlipNo = ""
                            rs3.Update
                        End If
                        rs4.Close
                        rs3.MoveNext
                    Loop
                End If
                rs3.Close
            End If
        End If
        rs2.Close
        
        rs1.MoveNext
    Loop
End If
rs1.Close
cn1.Close
MsgBox "Process completed successfully.", vbInformation + vbOKOnly
Exit Sub

ErrExit:
MsgBox Err.Description
End Sub

Private Sub RecreateDepositSlip(SlipId As String, InstituteCode As String)
On Error Resume Next
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

rs1.Open "Delete from DepositSlipDetails where SlipCode =" & SlipId & " and CentreCode ='" & InstituteCode & "'", cn1, adOpenDynamic, adLockPessimistic
    
SerStr = "Select CCChequeNo, CentreChequeDate, CCChequeAmt, ChqIDNo from ASPDC_DispatchSlipDetails ADSD inner join ASPDC_DispatchSlip ADS on adsd.dispatchslipcode= ads.dispatchslipcode " & _
         "where cmsdoneflag = 1 and CMSSlipNo ='" & SlipId & "' and misinstitutecode ='" & Left(InstituteCode, 3) & "' and liccode ='" & Right(InstituteCode, 2) & "'"

rs2.Open SerStr, cn1, adOpenDynamic, adLockReadOnly
If Not (rs2.BOF And rs2.EOF) Then
    
    
    rs1.Open "Select * from DepositSlipDetails where SlipCode =" & SlipId & " and CentreCode ='" & InstituteCode & "'", cn1, adOpenDynamic, adLockPessimistic
    rs2.MoveFirst
    Do While Not rs2.EOF
        If 1 = 1 Then
            'Check if chqidno exists in studentpayment
            rs3.Open "select sp.yearname, sp.rcptcode, sp.batchcode, sp.rollno, sp.SBEntryCode from studentpayment sp where sp.chequeidno ='" & rs2!ChqIdNo & "' and sp.recorddelflag =0", cn1, adOpenDynamic, adLockReadOnly
            If Not (rs3.BOF And rs3.EOF) Then
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
                rs1.Fields("ChequeIdNo").Value = rs2!ChqIdNo
                rs1.Fields("DownloadFlag").Value = "True"
                
                ChqAmt = ChqAmt + Val(rs2!CCChequeAmt)
                rs1.Update
                
            End If
            rs3.Close
        End If
        
        DoEvents
            
        rs2.MoveNext
    Loop
    rs1.Close
End If
rs2.Close

'Find Chq cnt and chq amt in aspdc_cms_log
rs1.Open "Select count(*) as ChqCnt, sum(CCChequeAmt) as SlipAmt from ASPDC_DispatchSlipDetails ADSD inner join ASPDC_DispatchSlip ADS on adsd.dispatchslipcode= ads.dispatchslipcode " & _
         "where cmsdoneflag = 1 and CMSSlipNo ='" & SlipId & "' and misinstitutecode ='" & Left(InstituteCode, 3) & "' and liccode ='" & Right(InstituteCode, 2) & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    ChqCnt = rs1!ChqCnt
    slipamt = rs1!slipamt
End If
rs1.Close

'Update these values in Depositslip and
rs2.Open "Select * from ASPDC_CMS_CentreLog where cmsslipno ='" & SlipId & "'", cn1, adOpenDynamic, adLockPessimistic
If (rs2.BOF And rs2.EOF) Then
    Dim SlipDate As Date
    rs1.Open "Select * from depositslip where slipno ='" & SlipId & "'", cn1, adOpenDynamic, adLockReadOnly
    If Not (rs1.BOF And rs1.EOF) Then
        SlipDate = rs1!SlipDate
    End If
    rs1.Close

    rs2.AddNew
    rs2!InstituteCode = Left(InstituteCode, 3)
    rs2!LicCode = Right(InstituteCode, 2)
    rs2!CMSDate = SlipDate
    rs2!CMSStatus = "SlipGenerated"
    rs2!FinalChequeCnt = ChqCnt
    rs2!FinalChequeAmt = slipamt
    rs2!CMSSlipNo = SlipId
    rs2.Update
Else
    'rs1.Open "Update ASPDC_CMS_CentreLog set FinalChequeCnt =" & ChqCnt & ", FinalChequeAmt =" & slipamt & " where cmsslipno ='" & SlipId & "'", cn1, adOpenDynamic, adLockPessimistic
    rs2!FinalChequeCnt = ChqCnt
    rs2!FinalChequeAmt = slipamt
    rs2.Update
End If
rs2.Close

'rs1.Open "Update depositslip set ChqCnt =" & ChqCnt & ", ChqAmt =" & slipamt & " where slipno ='" & SlipId & "'", cn1, adOpenDynamic, adLockPessimistic
rs1.Open "Select * from depositslip where slipno ='" & SlipId & "'", cn1, adOpenDynamic, adLockPessimistic
If Not (rs1.BOF And rs1.EOF) Then
    rs1!ChqCnt = ChqCnt
    rs1!ChqAmt = slipamt
    rs1.Update
End If
rs1.Close

cn1.Close
MsgBox "Process completed successfully.", vbInformation + vbOKOnly
End Sub

Private Sub mnuDeleteDispatchSlipEntry_Click()
On Error Resume Next
FrmDispatchSlip_Delete.Show vbModal
End Sub

Private Sub mnuDispatchMICRScan_Click()
FrmDispatch_MICREntry.Show
End Sub

Private Sub mnuDispatchSlipAuthorisation_Click()
On Error Resume Next
With FrmDispatchSlipAuthorisation
    .Show
    .WindowState = vbMaximized
    .ZOrder (0)
End With
End Sub

Private Sub mnuDispatchSlipCancellation_Click()
On Error Resume Next
With FrmDispatchSlip
    .cmdAdd.Visible = False
    .cmdCancelSlip.Visible = True
    .Show
    .WindowState = vbMaximized
    .ZOrder (0)
End With
End Sub

Private Sub mnuDispatchSlipChqEntry_Click()
On Error Resume Next
With FrmDispatchSlipChequeEntry
    .Show
    .WindowState = vbMaximized
    .ZOrder (0)
End With
End Sub

Private Sub mnuDispatchSlipEntry_Click()
On Error Resume Next
With FrmDispatchSlip
    .Show
    .WindowState = vbMaximized
    .ZOrder (0)
End With
End Sub

Private Sub mnuHoldChequeEntry_Click()
On Error Resume Next
With FrmHoldChequeEntry
    .Show
    .WindowState = vbMaximized
    .ZOrder (0)
End With
End Sub

Private Sub mnuInwardEntry_Click()
On Error Resume Next
If ModInit.LocationCode = "001" Then
    With FrmInwardEntryNew
        .Show
        .WindowState = vbMaximized
        .ZOrder (0)
    End With
Else
    With FrmInwardEntry
        .Show
        .WindowState = vbMaximized
        .ZOrder (0)
    End With
End If
End Sub

Private Sub mnuMakeCMS_Click()
On Error Resume Next
With FrmCMSDeposit_New ' FrmCMSDeposit_Sub4
    .Show
'    .WindowState = vbMaximized
    .ZOrder (0)
End With
End Sub

Private Sub mnuMonthlyStockCheck_Click()
With FrmMonthlyStockCheck
    .Show
    .WindowState = vbMaximized
    .ZOrder (0)
End With
End Sub

Private Sub mnuRepCMS_Click()
On Error Resume Next
'With FrmCMSDeposit
'    .Show
'    .WindowState = vbMaximized
'    .ZOrder (0)
'End With
End Sub

Private Sub mnuReturnChequeEntry_Click()
On Error Resume Next
With FrmReturnChequeEntry
    .Show
    .WindowState = vbMaximized
    .ZOrder (0)
End With
End Sub

Private Sub mnuReturnChequeRequests_Click()
On Error Resume Next
With FrmReturnChequeRequest
    .Show
    .WindowState = vbMaximized
    .ZOrder (0)
End With
End Sub

Private Sub mnuSearchCheque_Click()
FrmSearchCheque.Show
End Sub

Private Sub mnuSearchChequeNo_Click()
FrmSearchCheque_OnChqNo.Show
End Sub

Public Sub UpdateDashboard()
On Error Resume Next
If ModInit.LocationCode <> "001" Then Exit Sub

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

If Err.Number = -2147467259 Then
    Exit Sub
End If

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

rs1.Open "Select isnull(count(*),0) as RecCnt from ASPDC_DispatchSlipLogNew where (Location_Code ='" & ModInit.LocationCode & "' or Location_Code is Null) and DispatchSlipCode not in (Select DispatchSlipCode from ASPDC_DispatchSlip)", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    lblPending_Scan.Caption = Val("" & rs1!RecCnt)
End If
rs1.Close

rs1.Open "Select isnull(count(*),0) as RecCnt from ASPDC_DispatchSlip  where dispatchdate >='1 Jun 2014' and (Location_Code ='" & ModInit.LocationCode & "' or Location_Code is Null) and  ChequeCnt <> (select count(*) as ChqCnt from ASPDC_DispatchSlipDetails where DispatchSlipCode = ASPDC_DispatchSlip.DispatchSlipCode)", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    lblError_Scanning.Caption = Val("" & rs1!RecCnt)
End If
rs1.Close

rs1.Open "Select isnull(count(*),0) as RecCnt from ASPDC_DispatchSlip where (Location_Code ='" & ModInit.LocationCode & "' or Location_Code is Null) and SlipStatus = 1 and ChqEntryFlag =0 ", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    lblPending_Entry.Caption = Val("" & rs1!RecCnt)
End If
rs1.Close

rs1.Open "Select isnull(count(*),0) as RecCnt from ASPDC_DispatchSlip where (Location_Code ='" & ModInit.LocationCode & "' or Location_Code is Null) and SlipStatus = 1 and ChqEntryFlag =1 and AuthEntryFlag =0 and OpenChqCnt > 0 and CompleteEntryFlag =1 ", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    Me.lblPending_Auth.Caption = Val("" & rs1!RecCnt)
End If
rs1.Close

rs1.Open "Select isnull(count(*),0) as RecCnt from ASPDC_BounceChequeEntry where (Location_Code ='" & ModInit.LocationCode & "' or Location_Code is Null) and (MailSentFlag is Null or MailSentFlag = 0) and BounceEntryDate >='28 Jul 2014'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    Me.lblPending_BouncedMail.Caption = Val("" & rs1!RecCnt)
End If
rs1.Close

rs1.Open "select isnull(count(*),0) as RecCnt from ASPDC_DispatchSlip where slipentrydate >='18 Jan 2015' and isnull(OrderEngineUpdateFlag,0)  =0", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    Me.lblDispatchSlip.Caption = Val("" & rs1!RecCnt)
End If
rs1.Close

rs1.Open "select count(*) as RecCnt from aspdc_dispatchslipdetails where automapflag =0 and effectdownloadflag =0 and isnull(ccchqidno, '') <> ''", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    Me.lblCorrectedChq.Caption = Val("" & rs1!RecCnt)
End If
rs1.Close

rs1.Open "select count(*) as RecCnt from aspdc_dispatchslipdetails where returnflag =1 and returndate >='18 Jan 2015' and isnull(ReturnEffectDownloadFlag,0)  <=1", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    Me.lblRemovedChq.Caption = Val("" & rs1!RecCnt)
End If
rs1.Close

rs1.Open "select count(*) as RecCnt from ASPDC_BounceChequeEntry where bounceentrydate >='18 Jan 2015' and isnull(EffectDownloadFlag,0) <=1", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    Me.lblBounceCheque.Caption = Val("" & rs1!RecCnt)
End If
rs1.Close

'----- Added temporarily
'cn1.Close
'Exit Sub
'----- End Add

Dim CurTime As Integer
CurTime = TimeConvertor(Format(Time, "Hh:mm:AMPM"))

DoEvents

'If CurTime Mod 18 = 0 Then
    'Update SBEntryCode in order engine
'''''   MsgBox "In SBEntryCode", vbInformation
'''''   Timer1.Enabled = False
'''''    rs1.Open "select ChqIdNo from aspdc_dispatchslipdetails where cmssbentrycode is null and ccchequedate >='1 Dec 2014' and isnull(ReturnFlag,0) =0", cn1, adOpenDynamic, adLockReadOnly
'''''    If Not (rs1.BOF And rs1.EOF) Then
'''''        rs1.MoveFirst
'''''        Do While Not rs1.EOF
'''''            ModInit.WaitFlag = True
'''''
'''''
'''''
'''''            Dim APIStr4 As String
'''''            APIStr4 = "http://www.acountech.com/ImportStudentData_OrderEngine.aspx?DSC=" & rs1!ChqIdNo & "&LOC=001&Flag=1"
'''''
'''''            wbSend.Navigate2 APIStr4
'''''            Do While ModInit.WaitFlag = True
'''''                DoEvents
'''''            Loop
'''''
'''''
'''''            rs1.MoveNext
'''''        Loop
'''''    End If
'''''    rs1.Close
'''''    Timer1.Enabled = True
'''''    MsgBox "Out SBEntryCode", vbInformation
'End If

'If CurTime Mod 10 = 0 Then
    'Update Dispatch slip status in Order Engine
    rs1.Open "select * from ASPDC_DispatchSlip where slipentrydate >='18 Jan 2015' and isnull(OrderEngineUpdateFlag,0)  =0", cn1, adOpenDynamic, adLockReadOnly
    If Not (rs1.BOF And rs1.EOF) Then
        Timer1.Enabled = False
        ModInit.WaitFlag = True
        wbSend.Navigate2 "http://oe.mteducare.com/pdc_management/UpdateDispatchData_OrderEngine.aspx?Flag=5"
        Do While ModInit.WaitFlag = True
            DoEvents
        Loop
        Timer1.Enabled = True
    End If
    rs1.Close
'    cn1.Close
'    Exit Sub
'End If

'If CurTime Mod 12 = 0 Then
    'Update Corrected cheque details in order engine
    Dim Reason As String
    rs1.Open "select * from aspdc_dispatchslipdetails where automapflag =0 and effectdownloadflag =0 and isnull(ccchqidno, '') <> ''", cn1, adOpenDynamic, adLockReadOnly
    If Not (rs1.BOF And rs1.EOF) Then
        rs1.MoveFirst
        Timer1.Enabled = False
        Do While Not rs1.EOF
            ModInit.WaitFlag = True
            
            Reason = ""
            If rs1!CenterChequeNo = rs1!CCChequeNo And rs1!CenterchequeAmt <> rs1!CCChequeAmt Then
                Reason = "CR01"
            ElseIf rs1!CenterChequeNo <> rs1!CCChequeNo And rs1!CenterchequeAmt = rs1!CCChequeAmt Then
                Reason = "CR02"
            Else
                Reason = "CR03"
            End If
            
            Dim APIStr As String
            APIStr = "http://oe.mteducare.com/pdc_management/UpdateDispatchData_OrderEngine.aspx?Flag=2&P1=" & rs1!cmscenter_Code & "&P2=" & rs1!CMSSBEntryCode & "&P3=" & rs1!DispatchSlipCode & "&P4=" & rs1!DispatchSlipEntryCode & "&P5=" & rs1!ChqIdNo & "&P6=" & rs1!CenterChequeNo & "&P7=" & rs1!CentreChequeAmt & "&P8=" & rs1!CCChequeNo & "&P9=" & rs1!CCChequeAmt & "&P10=" & Reason
            
            wbSend.Navigate2 APIStr
            Do While ModInit.WaitFlag = True
                DoEvents
            Loop
            rs1.MoveNext
        Loop
        Timer1.Enabled = True
    End If
    rs1.Close
'    cn1.Close
'    Exit Sub
'End If

DoEvents
    
'If CurTime Mod 16 = 0 Then
    'Update Removed cheque details in order engine
    rs1.Open "select * from aspdc_dispatchslipdetails where returnflag =1 and returndate >='18 Jan 2015' and isnull(ReturnEffectDownloadFlag,0)  <=1", cn1, adOpenDynamic, adLockReadOnly
    If Not (rs1.BOF And rs1.EOF) Then
        rs1.MoveFirst
        Timer1.Enabled = False
        Do While Not rs1.EOF
            ModInit.WaitFlag = True
            
'            CR04 - Name Wrong on Cheque
'            CR05 - Figures in Word and Amount Mismatch
'            CR06 - Amount InCorrect

            'Admission Cancelled
            'Amount Mismatch            : CR05
            'Cheque Alteration / Correction
            'Cheque w/o Signature
            'Invalid Payee Name         : CR04
            'Non-CTS Cheque
            'Request by Centre
            'Stale Cheque
            'Wrong Cheque
            'Wrong Cheque Date

            'More reasons are yet to be added
            Reason = ""
            If "" & rs1!ReturnReason = "Amount Mismatch" Then
                Reason = "CR05"
            ElseIf "" & rs1!ReturnReason = "Invalid Payee Name" Then
                Reason = "CR04"
            Else
                Reason = "CR06"
            End If
            
            If Reason <> "" Then
            
                Dim APIStr1 As String
                APIStr1 = "http://oe.mteducare.com/pdc_management/UpdateDispatchData_OrderEngine.aspx?Flag=3&P1=" & rs1!cmscenter_Code & "&P2=" & rs1!CMSSBEntryCode & "&P3=" & rs1!DispatchSlipCode & "&P4=" & rs1!DispatchSlipEntryCode & "&P5=" & rs1!ChqIdNo & "&P6=" & rs1!CenterChequeNo & "&P7=" & rs1!CentreChequeAmt & "&P8=" & Reason
                
                wbSend.Navigate2 APIStr1
                Do While ModInit.WaitFlag = True
                    DoEvents
                Loop
            End If
            
            rs1.MoveNext
        Loop
        Timer1.Enabled = True
    End If
    rs1.Close
'    cn1.Close
'    Exit Sub
'End If
    
DoEvents

'If CurTime Mod 18 = 0 Then
   ' Update Bounced cheque details in order engine

    rs1.Open "select * from ASPDC_BounceChequeEntry AB inner join aspdc_dispatchslipdetails AD on AB.CCCHQIdNo = AD.CCCHQIdNo where bounceentrydate >='18 Jan 2015' and isnull(AB.EffectDownloadFlag,0) <=1", cn1, adOpenDynamic, adLockReadOnly
    'rs1.Open "select * from ASPDC_BounceChequeEntry where bounceentrydate >='18 Jan 2015' and isnull(EffectDownloadFlag,0) <=1", cn1, adOpenDynamic, adLockReadOnly
    If Not (rs1.BOF And rs1.EOF) Then
        rs1.MoveFirst
        Do While Not rs1.EOF
            ModInit.WaitFlag = True
            Timer1.Enabled = False

            'More reasons are yet to be added
            Reason = ""
            If "" & rs1!BouncePenaltyFlag = "1" Then
                Reason = "CR07"
            Else
                Reason = "CR08"
            End If

            If Reason <> "" Then

                Dim APIStr3 As String
                APIStr3 = "http://oe.mteducare.com/pdc_management/UpdateDispatchData_OrderEngine.aspx?Flag=4&P1=" & rs1!cmscenter_Code & "&P2=" & rs1!CMSSBEntryCode & "&P3=" & rs1!DispatchSlipCode & "&P4=" & rs1!DispatchSlipEntryCode & "&P5=" & rs1!ChqIdNo & "&P6=" & rs1!CenterChequeNo & "&P7=" & rs1!CentreChequeAmt & "&P8=" & Reason

                wbSend.Navigate2 APIStr3
                Do While ModInit.WaitFlag = True
                    DoEvents
                Loop
            End If

            rs1.MoveNext
        Loop
        Timer1.Enabled = True
    End If
    rs1.Close
'    cn1.Close
'    Exit Sub
'End If
    

    
    
    'Update CMS in order engine


cn1.Close
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim CurTime As Integer
CurTime = TimeConvertor(Format(Time, "Hh:mm:AMPM"))

If CurTime Mod 3 = 0 Then UpdateDashboard
End Sub



Public Function TimeConvertor(Tm As String) As Integer
'This function accepts time in HR:MN:AMPM format and converts
'it to a string format which can be directly checked for time
'calculations by following formula
'setp 1: Split time in three parts Hr, Mn, AmPm
'step 2: If AMPM Is PM Then Hr = Hr + 12
'step 3: HrInt = Hr * 60
'step 4: Result = HrInt + Mn
Dim HrInt As Long
Dim hR, mN, AMPM As String

'Step 1
Dim splitTM
splitTM = Split(Tm, ":")
hR = splitTM(0)
mN = splitTM(1)
AMPM = splitTM(2)

'Step 1
If (AMPM = "PM") And (Val(hR) <> 12) Then
    HrInt = Val(hR) + 12
ElseIf (AMPM = "PM") And (Val(hR) = 12) Then
    HrInt = Val(hR)
ElseIf (AMPM = "AM") And (Val(hR) <> 12) Then
    HrInt = Val(hR)
ElseIf (AMPM = "AM") And (Val(hR) = 12) Then
    HrInt = 0
End If

'Step 2
HrInt = HrInt * 60

'Step 3
TimeConvertor = HrInt + Val(mN)

End Function

Private Sub wbSend_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
ModInit.WaitFlag = False
'MsgBox wbSend.Document.Body.innerText
End Sub

