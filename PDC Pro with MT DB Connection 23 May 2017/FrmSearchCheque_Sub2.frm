VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmSearchCheque_Sub2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Cheque"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdMap 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Map"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox txtCCChequeNo 
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox txtCCChequeAmt 
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
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2730
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4815
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
      FormatString    =   "<Cheque ID|<Cheque No|<Cheque Date|<Centre Code|<Student Name|<Center Name"
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
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CC Cheque Number"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1665
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CC Cheque Amount"
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
      TabIndex        =   2
      Top             =   120
      Width           =   1650
   End
End
Attribute VB_Name = "FrmSearchCheque_Sub2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub SearchCheque()
On Error Resume Next
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

Grid.Rows = 1

'Check for duplicate barcode
rs1.Open "Select * from StudentPayment where chkNo ='" & Me.txtCCChequeNo.Text & "' and AmountPaid ='" & Format(txtCCChequeAmt.Text, "0.00") & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        '<Cheque ID|<Cheque No|<Cheque Date|<Centre Code|<Student Name|<Center Name
        Grid.Rows = Grid.Rows + 1
        Grid.TextMatrix(Grid.Rows - 1, 0) = rs1!ChequeIdNo
        Grid.TextMatrix(Grid.Rows - 1, 1) = rs1!ChkNo
        Grid.TextMatrix(Grid.Rows - 1, 2) = Format(rs1!ChkDate, "dd Mmm yyyy")
        Grid.TextMatrix(Grid.Rows - 1, 3) = rs1!InstituteCode & rs1!LicCode
        
        rs2.Open "Select * from tbl_mtmis_1 where sbentrycode ='" & rs1!SBEntryCode & "'", cn1, adOpenDynamic, adLockReadOnly
        If Not (rs2.BOF And rs2.EOF) Then
            Grid.TextMatrix(Grid.Rows - 1, 4) = rs2!StudentName
            Grid.TextMatrix(Grid.Rows - 1, 5) = rs2!CenterName
        End If
        rs2.Close
        
        rs1.MoveNext
    Loop
End If
rs1.Close
cn1.Close

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdMap_Click()
On Error Resume Next
If Grid.Rows = 1 Then Exit Sub

'Read ChequeIDNo
Dim ChqIdNo As String
ChqIdNo = Grid.TextMatrix(Grid.RowSel, 0)




'Update
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

'Check if same chqidno is mapped to any other cheque
rs1.Open "Select * from ASPDC_DispatchSlipDetails where ChqIDNo ='" & ChqIdNo & "' and CCChqIdNo <>'" & FrmSearchCheque.txtChqBarCode.Text & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    Dim Res As Integer
    Res = MsgBox("The same Cheque ID is mapped to another cheque.  Do you want to update the record?", vbQuestion + vbYesNo)
    
    If Res = 7 Then
    
        rs1.Close
        cn1.Close
        Exit Sub
    Else
        rs2.Open "Update ASPDC_DispatchSlipDetails set ChqIdNo ='' where  ChqIDNo ='" & ChqIdNo & "' and CCChqIdNo <>'" & FrmSearchCheque.txtChqBarCode.Text & "'", cn1, adOpenDynamic, adLockPessimistic
    End If
   
End If
rs1.Close

rs1.Open "Select * from ASPDC_DispatchSlipDetails where CCChqIdNo ='" & FrmSearchCheque.txtChqBarCode.Text & "'", cn1, adOpenDynamic, adLockPessimistic
If Not (rs1.BOF And rs1.EOF) Then
    rs1!ChqIdNo = ChqIdNo
    rs1.Update
End If
rs1.Close
cn1.Close

Unload Me

End Sub
