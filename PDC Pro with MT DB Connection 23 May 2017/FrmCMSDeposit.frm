VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmCMSDeposit 
   BackColor       =   &H00FFC0C0&
   Caption         =   "CMS Deposit Slip Entry"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13815
   Icon            =   "FrmCMSDeposit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6885
   ScaleWidth      =   13815
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdModifyCMS 
      Caption         =   "Modify CMS"
      Height          =   375
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdVerifyCMS 
      Caption         =   "Verify CMS"
      Height          =   375
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3930
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   6932
      _Version        =   393216
      Cols            =   7
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
      FormatString    =   "<Slip Number|<Company Name|<Centre Name|<Deposit Slip Date|<No of Instruments|>Total Slip Amount|<Verification Status"
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
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   240
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
      Format          =   125042691
      CurrentDate     =   39310
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   315
      Left            =   6120
      TabIndex        =   1
      Top             =   240
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
      Format          =   125042691
      CurrentDate     =   39310
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit Slip Date From"
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit Slip Date To"
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
      Left            =   4080
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "FrmCMSDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
On Error Resume Next
Unload Me
End Sub


Private Sub cmdModifyCMS_Click()
On Error Resume Next
If Grid.Rows = 1 Then Exit Sub

If Grid.TextMatrix(Grid.RowSel, 6) = "Verified" Then
    MsgBox "You can not Modify this CMS as it is already in Verified status.", vbCritical + vbOKOnly
    Exit Sub
End If

With FrmCMSDeposit_Modify
    .txtSlipNo.Text = Grid.TextMatrix(Grid.RowSel, 0)
    .FillGrid
    .Show vbModal
End With

End Sub

Private Sub cmdShow_Click()
ReadValues
End Sub

Private Sub cmdVerifyCMS_Click()
On Error Resume Next
FrmCMSDeposit_Verify.Show vbModal
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
'FillCompany
SetGrid
End Sub



Private Sub SetGrid()
On Error Resume Next
With Grid
    .Rows = 1
End With
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

Grid.Rows = 1

rs1.Open "select ac.institutecode, ac.liccode, cmsdate, CMSSlipNo, FinalChequeCnt, FinalChequeAmt , centrename, d.divisionname , ac.CMSVerifyFlag   from ASPDC_CMS_CentreLog AC inner join division D on ac.institutecode = d.institutecode left outer join g_centre_mis g on Ac.institutecode = g.institutecode and ac.liccode = g.liccode where cmsdate >='" & Format(dtFrom.Value, "dd Mmm yyyy") & "' and cmsdate <='" & Format(dtTo.Value, "dd Mmm yyyy") & "' and CMSStatus ='SlipGenerated' and FinalChequeCnt > 0 order by cmsdate, cmsslipno ", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        Grid.Rows = Grid.Rows + 1
        Grid.TextMatrix(Grid.Rows - 1, 0) = rs1.Fields("CMSSlipNo").Value
        Grid.TextMatrix(Grid.Rows - 1, 1) = rs1.Fields("divisionname").Value
        Grid.TextMatrix(Grid.Rows - 1, 2) = "" & rs1.Fields("centrename").Value
        Grid.TextMatrix(Grid.Rows - 1, 3) = Format(rs1.Fields("cmsdate").Value, "dd Mmm yyyy")
        Grid.TextMatrix(Grid.Rows - 1, 4) = rs1.Fields("FinalChequeCnt").Value
        Grid.TextMatrix(Grid.Rows - 1, 5) = rs1.Fields("FinalChequeAmt").Value
        If "" & rs1.Fields("CMSVerifyFlag").Value = "1" Then
            Grid.TextMatrix(Grid.Rows - 1, 6) = "Verified"
        Else
            Grid.TextMatrix(Grid.Rows - 1, 6) = "Not Verified"
        End If
        rs1.MoveNext
     Loop
End If
rs1.Close
cn1.Close

'Add Total
Grid.Rows = Grid.Rows + 1
Grid.TextMatrix(Grid.Rows - 1, 0) = "Total"
For Cnt = 1 To Grid.Rows - 2
    Grid.TextMatrix(Grid.Rows - 1, 4) = Val(Grid.TextMatrix(Grid.Rows - 1, 4)) + Val(Grid.TextMatrix(Cnt, 4))
    Grid.TextMatrix(Grid.Rows - 1, 5) = Val(Grid.TextMatrix(Grid.Rows - 1, 5)) + Val(Grid.TextMatrix(Cnt, 5))
Next

Exit Sub

ErrExit:
MsgBox Err.Description
End Sub

Private Sub SetGridWidth()
On Error Resume Next
Dim DefColWidth As Long
DefColWidth = (Grid.Width - 300) / 7
With Grid
    For Cnt = 0 To Grid.Cols - 1
        .ColWidth(Cnt) = DefColWidth
    Next
End With
End Sub

Private Sub Form_Resize()
On Error Resume Next
Grid.Width = Me.Width - Grid.Left - 240
Shape1.Width = Grid.Width
SetGridWidth
cmdClose.Left = Grid.Left + Grid.Width - cmdClose.Width
Grid.Height = Me.Height - Grid.Top - 1200
cmdClose.Top = Grid.Top + Grid.Height + 120
End Sub
