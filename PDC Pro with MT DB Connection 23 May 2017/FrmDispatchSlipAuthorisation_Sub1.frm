VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmDispatchSlipAuthorisation_Sub1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dispatch Slip Authorisation"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   11940
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRecheck 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Re-Check"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   360
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
      Left            =   5760
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
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox TxtUserName 
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7320
      Width           =   1695
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
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7320
      Width           =   2415
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
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Accept"
      Height          =   375
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7320
      Visible         =   0   'False
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
      Height          =   5490
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9684
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
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "<Cheque No|<Amount|<Cheque ID No|<Map Flag|<CC Cheque No|<CC Cheque Amount|<Entry Id|<Centre Cheque Date"
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
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Slip Entry Done By"
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
      TabIndex        =   14
      Top             =   7080
      Width           =   1605
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Open Items from Dispatch Slip"
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
      Top             =   1080
      Width           =   2580
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
      Left            =   5760
      TabIndex        =   12
      Top             =   240
      Width           =   1245
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Left            =   3000
      TabIndex        =   10
      Top             =   7080
      Width           =   1515
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dispatch Slip Number"
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
      Width           =   1875
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dispatch Date"
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
      Left            =   3000
      TabIndex        =   6
      Top             =   240
      Width           =   1245
   End
End
Attribute VB_Name = "FrmDispatchSlipAuthorisation_Sub1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub cmdCancel_Click()
On Error Resume Next
Unload Me
End Sub


Private Sub cmdRecheck_Click()
On Error Resume Next
Me.MousePointer = vbHourglass

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

rs1.Open "select * from ASPDC_DispatchSlipDetails where DispatchSlipCode ='" & Me.txtSlipNo.Text & "' and (CCChequeNo ='' or CCChequeNo is Null) order by DispatchSlipEntryCode", cn1, adOpenDynamic, adLockPessimistic
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        'Find if any entry exists with same chequeno and amout in open items
        rs2.Open "select * from ASPDC_DispatchSlip_OpenItems where DispatchSlipCode ='" & txtSlipNo.Text & "' and cast(CCChequeNo as bigint) =" & Val(rs1!CenterChequeNo) & " and  CCChequeAmt = " & Val(rs1!CentreChequeAmt) & " and  OpenEntryFlag = 1", cn1, adOpenDynamic, adLockPessimistic
        If Not (rs2.BOF And rs2.EOF) Then
            rs1!CCChequeNo = rs2!CCChequeNo
            rs1!CCChequeAmt = rs2!CCChequeAmt
            rs1!CCChequeDate = rs2!CCChequeDate
            rs1!CCCHQIdNo = rs2!CCCHQIdNo
            rs1!ManualMapReason = ""
            rs1!AutoMapFlag = 1
            rs1!EffectDownloadFlag = 1
            rs1.Update
       
            rs2!OpenEntryFlag = 0
            rs2.Update
        End If
        rs2.Close
        rs1.MoveNext
    Loop
End If
rs1.Close

'Find open item count
Dim OpenCnt As Integer
OpenCnt = 0
rs1.Open "select count(*) as OpenCnt from ASPDC_DispatchSlipDetails where DispatchSlipCode ='" & Me.txtSlipNo.Text & "' and (CCChequeNo ='' or CCChequeNo is Null)", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    OpenCnt = Val("" & rs1!OpenCnt)
End If
rs1.Close

rs1.Open "Update ASPDC_DispatchSlip set OpenChqCnt =" & OpenCnt & " where DispatchSlipCode ='" & Me.txtSlipNo.Text & "'", cn1, adOpenDynamic, adLockPessimistic


cn1.Close

ReadOpenItems
MsgBox "Re-Checking activity successfully finished.", vbInformation + vbOKOnly
End Sub

Private Sub cmdSave_Click()
On Error Resume Next

If checkvalid = False Then Exit Sub

Me.MousePointer = vbHourglass

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

'Start mapping the two entries
Dim str As String
'str = "Select * from ASPDC_DispatchSlipDetails where DispatchSlipCode ='" & Trim(UCase(Me.txtSlipNo.Text)) & "'"

Dim AutoChqCnt As Integer
AutoChqCnt = 0

'rs1.Open str, cn1, adOpenDynamic, adLockPessimistic
'If Not (rs1.BOF And rs1.EOF) Then
'    rs1.MoveFirst
'    Do While Not rs1.EOF
'        'Check if this cheque is already present in grid
'        For Cnt = 1 To Grid.Rows - 1
'            If Val(Grid.TextMatrix(Cnt, 0)) = Val(rs1!CenterChequeNo) And Val(Grid.TextMatrix(Cnt, 0)) <> 0 And Val(Grid.TextMatrix(Cnt, 2)) = rs1!CentreChequeAmt And Grid.TextMatrix(Cnt, 4) = "" Then
'                Grid.TextMatrix(Cnt, 4) = "Mapped"
'                Grid.TextMatrix(Cnt, 5) = rs1!ChqIDNo
'                AutoChqCnt = AutoChqCnt + 1
'
'                rs1!AutoMapFlag = 1
'                rs1!CCChequeNo = rs1!CenterChequeNo
'                rs1!CCChequeAmt = rs1!CentreChequeAmt
'                rs1!CCChequeDate = DateValue(Left(Grid.TextMatrix(Cnt, 1), 2) & "-" & MonthName(Mid(Grid.TextMatrix(Cnt, 1), 3, 2)) & "-" & Right(Grid.TextMatrix(Cnt, 1), 4))
'                rs1!EffectDownloadFlag = 1  'No need to download the effect at centre as it is automapped
'                rs1.Update
'                Exit For
'            End If
'        Next
'        rs1.MoveNext
'    Loop
'End If
'rs1.Close

'Store open entries in ASPDC_DispatchSlip_OpenItems
str = "Select * from ASPDC_DispatchSlip_OpenItems where DispatchSlipCode ='" & Trim(UCase(Me.txtSlipNo.Text)) & "'"
rs1.Open str, cn1, adOpenDynamic, adLockPessimistic
For Cnt = 1 To Grid.Rows - 1
    'If Grid.TextMatrix(Cnt, 4) = "" Then
        rs1.AddNew
        rs1.Fields("DispatchSlipCode").Value = Trim(txtSlipNo.Text)
        rs1!OpenItemEntryNo = Cnt
        rs1!CCChequeNo = rs1!CenterChequeNo
        rs1!CCChequeAmt = rs1!CentreChequeAmt
        rs1!CCChequeDate = DateValue(Left(Grid.TextMatrix(Cnt, 1), 2) & "-" & MonthName(Mid(Grid.TextMatrix(Cnt, 1), 3, 2)) & "-" & Right(Grid.TextMatrix(Cnt, 1), 4))
        rs1!OpenEntryFlag = 1
        rs1!LinkedCHQIdNo = ""
        rs1.Update
    'End If
Next
rs1.Close
    
rs1.Open "Select * from ASPDC_DispatchSlip where DispatchSlipCode ='" & txtSlipNo.Text & "'", cn1, adOpenDynamic, adLockPessimistic
If Not (rs1.BOF And rs1.EOF) Then
    rs1.Fields("ChqEntryFlag").Value = 1
    rs1.Fields("ChqEntryUserName").Value = TxtUserName.Text
    rs1.Fields("ChqEntryDate").Value = Date

    rs1.Fields("AuthEntryFlag").Value = 0
    rs1.Fields("AuthEntryUserName").Value = ""
    rs1.Fields("AuthEntryDate").Value = ""
    
    rs1.Fields("AutoMapChqCnt").Value = AutoChqCnt
    rs1.Fields("ManualMapChqCnt").Value = 0
    rs1.Fields("OpenChqCnt").Value = Val(Me.txtChequeCnt.Text) - AutoChqCnt
    rs1.Update
End If
rs1.Close
cn1.Close

Me.MousePointer = 0

MsgBox "Dispatch Slip - Cheque Entry successfully done.", vbOKOnly + vbInformation
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
    .ColWidth(6) = .ColWidth(0)
    .ColWidth(7) = .ColWidth(0)
    .Rows = 1
End With
End Sub

Public Sub ReadOpenItems()
On Error GoTo ErrExit
Grid.Rows = 1
Me.MousePointer = vbHourglass

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

rs1.Open "select * from ASPDC_DispatchSlipDetails where DispatchSlipCode ='" & Me.txtSlipNo.Text & "' and (CCChequeNo ='' or CCChequeNo is Null) order by DispatchSlipEntryCode", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        Grid.Rows = Grid.Rows + 1
        Grid.TextMatrix(Grid.Rows - 1, 0) = rs1!CenterChequeNo
        Grid.TextMatrix(Grid.Rows - 1, 1) = rs1!CentreChequeAmt
        Grid.TextMatrix(Grid.Rows - 1, 2) = rs1!ChqIDNo
        Grid.TextMatrix(Grid.Rows - 1, 3) = ""
        Grid.TextMatrix(Grid.Rows - 1, 4) = ""
        Grid.TextMatrix(Grid.Rows - 1, 5) = ""
        Grid.TextMatrix(Grid.Rows - 1, 6) = ""
        Grid.TextMatrix(Grid.Rows - 1, 7) = Format(rs1!CentreChequeDate, "dd Mmm yyyy")
        rs1.MoveNext
    Loop
End If
rs1.Close
cn1.Close
Me.MousePointer = 0
Exit Sub

ErrExit:
Me.MousePointer = 0
'cn1.Close
End Sub

Private Function checkvalid() As Boolean
On Error Resume Next
If Trim(Me.txtSlipNo.Text) = "" Then
    MsgBox "Enter Dispatch Slip Number.", vbInformation + vbOKOnly
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

With FrmDispatchSlipAuthorisation_Sub2
    .txtSlipNo.Text = Me.txtSlipNo.Text
    .txtChequeEntry.Text = Grid.TextMatrix(Grid.RowSel, 0)
    .txtChequeAmt.Text = Grid.TextMatrix(Grid.RowSel, 1)
    .txtChqIdNo.Text = Grid.TextMatrix(Grid.RowSel, 2)
    .TxtUserName.Text = FrmDispatchSlipAuthorisation.Grid.TextMatrix(FrmDispatchSlipAuthorisation.Grid.RowSel, 10)
    .ReadOpenItems
    .Show vbModal
End With
End Sub
