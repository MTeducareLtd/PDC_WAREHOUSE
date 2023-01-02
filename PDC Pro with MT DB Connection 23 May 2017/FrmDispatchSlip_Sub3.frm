VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmDispatchSlip_Sub3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search from Inwarded Slips"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Accept"
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4530
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7990
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
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "<Inward Number|<Deposit Slip Number|<Deposit Slip Date|<Inward Date|<Cheque Cnt|<Slip Type"
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
End
Attribute VB_Name = "FrmDispatchSlip_Sub3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo ErrExit
Grid.Rows = 1

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

If Err.Number = -2147467259 Then
    MsgBox "Unable to connect to server.", vbCritical + vbOKOnly
    Exit Sub
End If

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset


Me.MousePointer = vbHourglass

rs1.Open "Select * from ASPDC_DispatchSlipLogNew where (Location_Code ='" & ModInit.LocationCode & "' or Location_Code is Null) and DispatchSlipCode not in (Select DispatchSlipCode from ASPDC_DispatchSlip) order by InwardNo", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        Grid.Rows = Grid.Rows + 1
        Grid.TextMatrix(Grid.Rows - 1, 0) = rs1.Fields("InwardNo").Value
        Grid.TextMatrix(Grid.Rows - 1, 1) = rs1.Fields("DispatchSlipCode").Value
        Grid.TextMatrix(Grid.Rows - 1, 2) = Format(rs1.Fields("DispatchDate").Value, "dd Mmm yyyy")
        Grid.TextMatrix(Grid.Rows - 1, 3) = Format(rs1.Fields("InwardDate").Value, "dd Mmm yyyy")

        Grid.TextMatrix(Grid.Rows - 1, 4) = "" & rs1.Fields("ChequeCnt").Value
        Grid.TextMatrix(Grid.Rows - 1, 5) = "" & rs1.Fields("SlipType").Value

        
        rs1.MoveNext
     Loop
End If
rs1.Close
cn1.Close
Me.MousePointer = 0

Exit Sub

ErrExit:
Me.MousePointer = 0
MsgBox "Error : " & Err.Description
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
If Grid.Rows = 1 Then Exit Sub

With FrmDispatchSlip_Sub1
    .txtSlipNo.Text = Grid.TextMatrix(Grid.RowSel, 1)
    .txtChequeCnt.Text = Grid.TextMatrix(Grid.RowSel, 4)
    .txtSlipDate.Text = Format(Grid.TextMatrix(Grid.RowSel, 2), "ddmmyyyy")
    
    .txtSlipNo.Locked = True
    .txtSlipDate.Locked = True
    .txtChequeCnt.Locked = True
    
    .txtChequeEntry.SetFocus
    .Show vbModal
End With

cmdRefresh_Click
End Sub

Private Sub Form_Load()
On Error Resume Next
cmdRefresh_Click

With Grid
    For Cnt = 0 To .Cols - 1
        .ColWidth(Cnt) = (.Width - 350) / (.Cols)
    Next
    
    .ColWidth(1) = .ColWidth(0) + 500
    .ColWidth(4) = .ColWidth(0) - 500
End With

End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cmdSave_Click
    KeyAscii = 0
End If
End Sub
