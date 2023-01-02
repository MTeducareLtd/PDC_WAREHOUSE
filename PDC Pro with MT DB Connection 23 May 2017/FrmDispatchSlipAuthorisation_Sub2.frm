VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmDispatchSlipAuthorisation_Sub2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dispatch Slip Authorisation Entry"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9420
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optAllChq 
      Caption         =   "All &Open Items"
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
      Left            =   5760
      TabIndex        =   2
      Top             =   1200
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.OptionButton optChqAmt 
      Caption         =   "Matching Cheque &Amount"
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
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.OptionButton optChqNo 
      Caption         =   "Matching Cheque &Number"
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
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   2535
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5640
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   1695
   End
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
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   855
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
      Left            =   5550
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtChequeEntry 
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
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtChequeAmt 
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3210
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5662
      _Version        =   393216
      Cols            =   6
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
      FormatString    =   "<CC Cheque No|<CC Cheque Amount|<CC Cheque Date|<Type|<Entry Id|<Cheque Barcode"
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
   Begin VB.Label lblStudentName 
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
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Entry Done By"
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
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label2 
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
      Left            =   5520
      TabIndex        =   11
      Top             =   240
      Width           =   1875
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Centre Cheque Number"
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
      Top             =   240
      Width           =   1980
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Centre Cheque Amount"
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
      TabIndex        =   8
      Top             =   240
      Width           =   1965
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select a Open Item from Cheque Entry"
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
      Top             =   1680
      Width           =   3270
   End
End
Attribute VB_Name = "FrmDispatchSlipAuthorisation_Sub2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MisMatchCloseFlag As Boolean

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

If Me.optChqNo.Value = True Then
    rs1.Open "select * from ASPDC_DispatchSlip_OpenItems where DispatchSlipCode ='" & FrmDispatchSlipAuthorisation_Sub1.txtSlipNo.Text & "' and cast(CCChequeNo as bigint) =" & Val(Me.txtChequeEntry.Text) & " and  OpenEntryFlag = 1 order by OpenItemEntryNo", cn1, adOpenDynamic, adLockReadOnly
ElseIf Me.optAllChq.Value = True Then
    rs1.Open "select * from ASPDC_DispatchSlip_OpenItems where DispatchSlipCode ='" & FrmDispatchSlipAuthorisation_Sub1.txtSlipNo.Text & "' and OpenEntryFlag = 1 order by OpenItemEntryNo", cn1, adOpenDynamic, adLockReadOnly
ElseIf Me.optChqAmt.Value = True Then
    rs1.Open "select * from ASPDC_DispatchSlip_OpenItems where DispatchSlipCode ='" & FrmDispatchSlipAuthorisation_Sub1.txtSlipNo.Text & "' and CCChequeAmt = " & Val(Me.txtChequeAmt.Text) & " and OpenEntryFlag = 1 order by OpenItemEntryNo", cn1, adOpenDynamic, adLockReadOnly
End If

If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        Grid.Rows = Grid.Rows + 1
        Grid.TextMatrix(Grid.Rows - 1, 0) = rs1!CCChequeNo
        Grid.TextMatrix(Grid.Rows - 1, 1) = rs1!CCChequeAmt
        Grid.TextMatrix(Grid.Rows - 1, 2) = Format(rs1!CCChequeDate, "dd Mmm yyyy")
        Grid.TextMatrix(Grid.Rows - 1, 3) = rs1!CCChequeType
        Grid.TextMatrix(Grid.Rows - 1, 4) = rs1!OpenItemEntryNo
        Grid.TextMatrix(Grid.Rows - 1, 5) = rs1!CCCHQIdNo
        rs1.MoveNext
    Loop
End If
rs1.Close
cn1.Close
Me.MousePointer = 0
Exit Sub

ErrExit:
On Error Resume Next
Me.MousePointer = 0
'cn1.Close
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
With Grid
    .Rows = 1
    For Cnt = 0 To .Cols - 1
        .ColWidth(Cnt) = (.Width - 350) / (.Cols)
    Next
End With
End Sub

Private Sub Grid_DblClick()
On Error Resume Next
If Grid.Rows = 1 Then
    Dim res As Integer
    res = MsgBox("Open cheque entry does not exits.  Do you want to add a new entry.", vbQuestion + vbYesNo)
    If res = 7 Then Exit Sub
End If

With FrmDispatchSlipAuthorisation_Sub3
    MisMatchCloseFlag = False
    
    .txtSlipNo.Text = txtSlipNo.Text
    .txtChequeEntry.Text = Me.txtChequeEntry.Text
    .txtChequeAmt.Text = Me.txtChequeAmt.Text
    .txtChqIdNo.Text = Me.txtChqIdNo.Text
    If Grid.Rows > 1 Then
        .txtCCChequeNo.Text = Grid.TextMatrix(Grid.RowSel, 0)
        .txtCCChequeAmt.Text = Grid.TextMatrix(Grid.RowSel, 1)
        .dtChqDate.Value = Grid.TextMatrix(Grid.RowSel, 2)
        .cboChqType.Text = Grid.TextMatrix(Grid.RowSel, 3)
    
        .txtOpenItemNo.Text = Grid.TextMatrix(Grid.RowSel, 4)
        .txtChqBarCode.Text = Grid.TextMatrix(Grid.RowSel, 5)
        .txtChqBarCode.Enabled = False
    Else
        .txtChqBarCode.Enabled = True
        .txtChqBarCode.Locked = False
    End If
    
    If .txtChequeEntry.Text <> .txtCCChequeNo.Text And Val(.txtChequeAmt.Text) <> Val(.txtCCChequeAmt.Text) Then
        .chkReturnToCentre.Value = vbChecked
        If ModInit.PDCUserName = "Mithun" Then
            .chkReturnToCentre.Enabled = True
        Else
            .chkReturnToCentre.Enabled = False
        End If
    Else
        .chkReturnToCentre.Value = vbUnchecked
        If ModInit.PDCUserName = "Mithun" Then
            .chkReturnToCentre.Enabled = True
        Else
            .chkReturnToCentre.Enabled = False
        End If
    End If
    
    .Show vbModal
    
    If MisMatchCloseFlag = True Then
        FrmDispatchSlipAuthorisation_Sub1.ReadOpenItems
        Unload Me
    End If
End With
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    Grid_DblClick
    KeyAscii = 0
End If
End Sub

Private Sub optAllChq_Click()
ReadOpenItems
End Sub

Private Sub optChqAmt_Click()
ReadOpenItems
End Sub

Private Sub optChqNo_Click()
ReadOpenItems
End Sub
