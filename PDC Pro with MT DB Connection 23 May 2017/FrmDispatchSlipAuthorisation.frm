VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmDispatchSlipAuthorisation 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Dispatch Slip Authorisation"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   12660
   WindowState     =   2  'Maximized
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
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Start Entry"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2730
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   4815
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
      FormatString    =   $"FrmDispatchSlipAuthorisation.frx":0000
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
Attribute VB_Name = "FrmDispatchSlipAuthorisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdAdd_Click()
On Error Resume Next
If Grid.Rows = 1 Then Exit Sub

With FrmDispatchSlipAuthorisation_Sub1
    .txtSlipNo.Text = Grid.TextMatrix(Grid.RowSel, 0)
    .txtSlipDate.Text = Grid.TextMatrix(Grid.RowSel, 1)
    .txtChequeCnt.Text = Grid.TextMatrix(Grid.RowSel, 5)
    .TxtUserName.Text = Grid.TextMatrix(Grid.RowSel, 8)
    .ReadOpenItems
    .Show vbModal
End With
ReadValues
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
Unload Me
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
End Sub

Private Sub ReadValues()
On Error GoTo ErrExit
Me.MousePointer = vbHourglass

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset


Grid.Rows = 1
rs1.Open "Select DispatchSlipCode, DispatchDate, new_institutecode as MISInstituteCode, a.new_liccode as LicCode, g.Source_Center_Name as CentreName, OpenChqCnt, ChequeCnt, ChqEntryDate, ChqEntryUserName, SlipStatus, SlipEntryDate, SlipEntryUserName, d.Source_Division_ShortDesc DivisionName from ASPDC_DispatchSlip A inner join C008_Centers g on a.new_InstituteCode + a.new_LicCode = g.Source_Center_Code left outer join C006_Division D on a.new_InstituteCode = d.source_division_code where SlipStatus = 1 and ChqEntryFlag =1 and AuthEntryFlag =0 and OpenChqCnt > 0 and CompleteEntryFlag =1 and a.Location_Code ='" & ModInit.LocationCode & "' order by DispatchDate, DispatchSlipCode ", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        Grid.Rows = Grid.Rows + 1
        Grid.TextMatrix(Grid.Rows - 1, 0) = rs1.Fields("DispatchSlipCode").Value
        Grid.TextMatrix(Grid.Rows - 1, 1) = Format(rs1.Fields("DispatchDate").Value, "dd Mmm yyyy")
        Grid.TextMatrix(Grid.Rows - 1, 2) = rs1!MISInstituteCode & rs1!LicCode
        Grid.TextMatrix(Grid.Rows - 1, 3) = rs1!DivisionName
        Grid.TextMatrix(Grid.Rows - 1, 4) = rs1!CentreName
        Grid.TextMatrix(Grid.Rows - 1, 5) = rs1.Fields("OpenChqCnt").Value
        If rs1!SlipStatus = 1 Then
            Grid.TextMatrix(Grid.Rows - 1, 6) = "Accepted"
        Else
            Grid.TextMatrix(Grid.Rows - 1, 6) = "Cancelled"
        End If
        Grid.TextMatrix(Grid.Rows - 1, 7) = Format(rs1.Fields("SlipEntryDate").Value, "dd Mmm yyyy")
        Grid.TextMatrix(Grid.Rows - 1, 8) = rs1.Fields("SlipEntryUserName").Value
        Grid.TextMatrix(Grid.Rows - 1, 9) = Format(rs1.Fields("ChqEntryDate").Value, "dd Mmm yyyy")
        Grid.TextMatrix(Grid.Rows - 1, 10) = rs1.Fields("ChqEntryUserName").Value
        
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
Grid.Height = Me.Height - Grid.Top - 1200
cmdClose.Top = Grid.Top + Grid.Height + 120
End Sub


