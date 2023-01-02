VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmSearchCheque_OnChqNo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Cheque on Cheque No"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Search"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1695
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
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   2415
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
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3450
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   6085
      _Version        =   393216
      Cols            =   14
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
      FormatString    =   $"FrmSearchCheque_OnChqNo.frx":0000
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
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   1650
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1665
   End
End
Attribute VB_Name = "FrmSearchCheque_OnChqNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
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

'Check for duplicate barcode
rs1.Open "Select * from ASPDC_DispatchSlipDetails where CCChequeNo ='" & Me.txtCCChequeNo.Text & "' and CCChequeAmt =" & Me.txtCCChequeAmt.Text & "", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        '<Barcode|<Cheque Date|<CMS Status|<CMS Slip No|<CMS Slip Date|<Centre Code|<Student Name
        '<Barcode|<Cheque Date|<CMS Status|<CMS Slip No|<CMS Slip Date|<Centre Code|<Student Name|<Return Status|<Return Date|<SBEntryCode|<Cheque Entry Date|<Bounced Status|<Bounced Entry Date|<Cheque Location
        Grid.Rows = Grid.Rows + 1
        Grid.TextMatrix(Grid.Rows - 1, 0) = rs1!CCChqIdNo
        Grid.TextMatrix(Grid.Rows - 1, 1) = Format(rs1!CCChequeDate, "dd Mmm yyyy")

        If rs1!CMSDoneFlag = 1 Then
            Grid.TextMatrix(Grid.Rows - 1, 2) = "CMS Done"
            Grid.TextMatrix(Grid.Rows - 1, 3) = rs1!CMSSlipNo
            Grid.TextMatrix(Grid.Rows - 1, 4) = Format(rs1!CMSSlipDate, "dd Mmm yyyy")
            Grid.TextMatrix(Grid.Rows - 1, 9) = "" & rs1!CMSSBEntryCode
        Else
            Grid.TextMatrix(Grid.Rows - 1, 2) = ""
            Grid.TextMatrix(Grid.Rows - 1, 3) = ""
            Grid.TextMatrix(Grid.Rows - 1, 4) = ""
            Grid.TextMatrix(Grid.Rows - 1, 9) = ""
        End If
        Grid.TextMatrix(Grid.Rows - 1, 5) = rs1!DispatchSlipCode

'        rs2.Open "Select studentname, currentstudentlflag, T.SBentryCode from Tbl_Mtmis_1 T inner join StudentPayment SP on T.Sbentrycode = sp.sbentrycode where sp.ChequeIdNo ='" & rs1!ChqIdNo & "'", cn1, adOpenDynamic, adLockReadOnly
'        If Not (rs2.BOF And rs2.EOF) Then
'            If rs2!currentstudentlflag = "0" Then
'                Grid.TextMatrix(Grid.Rows - 1, 6) = "X-" & rs2!StudentName
'            Else
'
'                Grid.TextMatrix(Grid.Rows - 1, 6) = rs2!StudentName
'            End If
'            If Grid.TextMatrix(Grid.Rows - 1, 9) = "" Then
'                Grid.TextMatrix(Grid.Rows - 1, 9) = "" & rs2!SBEntryCode
'            End If
'        Else
'            Grid.TextMatrix(Grid.Rows - 1, 6) = "No matching student record found"
'        End If

        Grid.TextMatrix(Grid.Rows - 1, 6) = "" & rs1!StudentName
        If rs1!ReturnFlag = 1 Then
            Grid.TextMatrix(Grid.Rows - 1, 7) = "Returned to Centre"
            Grid.TextMatrix(Grid.Rows - 1, 8) = Format("" & rs1!ReturnDate, "dd Mmm yyyy")
        Else
            Grid.TextMatrix(Grid.Rows - 1, 7) = ""
            Grid.TextMatrix(Grid.Rows - 1, 8) = ""
        End If
'        rs2.Close
        
        rs2.Open "Select ChqEntryDate from ASPDC_DispatchSlip where DispatchSlipCode ='" & rs1!DispatchSlipCode & "'", cn1, adOpenDynamic, adLockReadOnly
        If Not (rs2.BOF And rs2.EOF) Then
            Grid.TextMatrix(Grid.Rows - 1, 10) = Format(rs2!ChqEntryDate, "dd Mmm yyyy")
        End If
        rs2.Close

        If rs1!CMSDoneFlag = 1 Then
            rs2.Open "Select * from ASPDC_BounceChequeEntry where CCChqIDNo ='" & rs1!CCChqIdNo & "'", cn1, adOpenDynamic, adLockReadOnly
            If Not (rs2.BOF And rs2.EOF) Then
                Grid.TextMatrix(Grid.Rows - 1, 11) = "Bounced"
                Grid.TextMatrix(Grid.Rows - 1, 12) = Format(rs2!BounceEntryDate, "dd Mmm yyyy")
                Grid.TextMatrix(Grid.Rows - 1, 13) = "" & rs2!chequelocation
            End If
            rs2.Close

        End If
        
        rs1.MoveNext
    Loop

Else
    MsgBox "Cheque Barcode not found.", vbInformation + vbOKOnly
End If
rs1.Close


Exit Sub

ErrExit:
MsgBox Err.Description
End Sub
