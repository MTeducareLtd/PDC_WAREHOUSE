VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSearchCheque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Cheque"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMap 
      Caption         =   "Map"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmsAccept 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Accept"
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtCentreCode 
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox txtStudentName 
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
      TabIndex        =   20
      Top             =   3240
      Width           =   5055
   End
   Begin VB.TextBox txtCMSDate 
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
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox txtCMSNo 
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox txtCMSFlag 
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
      TabIndex        =   14
      Top             =   2400
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
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1560
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
      TabIndex        =   6
      Top             =   1560
      Width           =   2415
   End
   Begin VB.ComboBox cboChqType 
      Enabled         =   0   'False
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
      ItemData        =   "FrmSearchCheque.frx":0000
      Left            =   240
      List            =   "FrmSearchCheque.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Search"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtChqBarCode 
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
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker dtChqDate 
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
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
      Format          =   103219203
      CurrentDate     =   41403
   End
   Begin MSComCtl2.DTPicker dtCentreDate 
      Height          =   375
      Left            =   7920
      TabIndex        =   12
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
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
      Format          =   103219203
      CurrentDate     =   41403
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Centre Code"
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
      Left            =   5400
      TabIndex        =   23
      Top             =   3000
      Width           =   1065
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Student Name"
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
      TabIndex        =   21
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CMS Date"
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
      Left            =   7920
      TabIndex        =   19
      Top             =   2160
      Width           =   870
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CMS Number"
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
      Left            =   5400
      TabIndex        =   17
      Top             =   2160
      Width           =   1110
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CMS Status"
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
      TabIndex        =   15
      Top             =   2160
      Width           =   1005
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Centre Cheque Date"
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
      Left            =   7920
      TabIndex        =   13
      Top             =   1320
      Width           =   1740
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
      TabIndex        =   11
      Top             =   1320
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
      TabIndex        =   10
      Top             =   1320
      Width           =   1650
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CC Cheque Date"
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
      Left            =   5400
      TabIndex        =   9
      Top             =   1320
      Width           =   1425
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CC Cheque Type"
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
      TabIndex        =   8
      Top             =   2160
      Width           =   1440
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Barcode"
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
      TabIndex        =   1
      Top             =   240
      Width           =   1425
   End
End
Attribute VB_Name = "FrmSearchCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdMap_Click()
On Error Resume Next
With FrmSearchCheque_Sub2
    .txtCCChequeNo.Text = Me.txtCCChequeNo.Text
    .txtCCChequeAmt.Text = Me.txtCCChequeAmt.Text
    .SearchCheque
    .Show vbModal
End With
End Sub

Public Sub cmdSave_Click()
On Error GoTo ErrExit
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

'Check for duplicate barcode
rs1.Open "Select * from ASPDC_DispatchSlipDetails where CCChqIdNo ='" & Me.txtChqBarCode.Text & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    Me.txtCCChequeNo.Text = rs1!CCChequeNo
    Me.txtCCChequeAmt.Text = rs1!CCChequeAmt
    Me.dtCentreDate.Value = rs1!CentreChequeDate
    Me.dtChqDate.Value = rs1!CCChequeDate
    Me.txtStudentName.Text = "" & rs1!StudentName
    
    If "" & rs1!CCChequeType <> "" Then
        Me.cboChqType.Text = "" & rs1!CCChequeType
    Else
        Me.cboChqType.ListIndex = 0
    End If
    
    If rs1!CMSDoneFlag = 1 Then
        Me.txtCMSFlag.Text = "CMS Done"
        Me.txtCMSNo.Text = rs1!CMSSlipNo
        Me.txtCMSDate.Text = Format(rs1!CMSSlipDate, "dd Mmm yyyy")
        Me.cmsAccept.Enabled = False
        
       
    Else
        Me.txtCMSFlag.Text = ""
        Me.txtCMSNo.Text = ""
        Me.txtCMSDate.Text = ""
        Me.cmsAccept.Enabled = True
        'cmdShowDupli.Visible = False
        
        
    End If
    Me.txtCentreCode.Text = rs1!DispatchSlipCode
    
'    rs2.Open "Select studentname, currentstudentlflag from Tbl_Mtmis_1 T inner join StudentPayment SP on T.Sbentrycode = sp.sbentrycode where sp.ChequeIdNo ='" & rs1!ChqIdNo & "'", cn1, adOpenDynamic, adLockReadOnly
'    If Not (rs2.BOF And rs2.EOF) Then
'        If rs2!currentstudentlflag = "0" Then
'            Me.txtStudentName.Text = "X-" & rs2!StudentName
'        Else
'
'            Me.txtStudentName.Text = rs2!StudentName
'        End If
'    Else
'        txtStudentName.Text = "No matching student record found"
'    End If
'    rs2.Close
    
    
    
Else
    MsgBox "Cheque Barcode not found.", vbInformation + vbOKOnly
End If
rs1.Close


Exit Sub

ErrExit:
MsgBox Err.Description
End Sub

Private Sub cmdShowDupli_Click()
On Error Resume Next
With FrmSearchCheque_Sub1
    .txtChqBarCode.Text = Me.txtChqBarCode.Text
    .FillDuplicateCMS
    .Show vbModal
End With
End Sub

Private Sub cmsAccept_Click()
On Error Resume Next
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

'Check for duplicate barcode
rs1.Open "Select * from ASPDC_DispatchSlipDetails where CCChqIdNo ='" & Me.txtChqBarCode.Text & "'", cn1, adOpenDynamic, adLockPessimistic
If Not (rs1.BOF And rs1.EOF) Then
    If rs1!CentreChequeDate <> Me.dtCentreDate.Value Then
        rs1!OGCentreChequeDate = rs1!CentreChequeDate
        rs1!CentreChequeDate = Me.dtCentreDate.Value
        rs1!CentreChequeDateChangeFlag = 1
    End If
    rs1!CCChequeDate = Me.dtChqDate.Value
    rs1.Update
End If
rs1.Close
cn1.Close

MsgBox "Cheque Date Changed successfully.", vbInformation + vbOKOnly
End Sub

Private Sub Form_Load()
On Error Resume Next
If ModInit.PDCUserType = 1 Then 'Manager
    Me.dtCentreDate.Enabled = True
    Me.dtChqDate.Enabled = True
    Me.cmsAccept.Visible = True
Else
    Me.dtCentreDate.Enabled = False
    Me.dtChqDate.Enabled = False
    Me.cmsAccept.Visible = False
End If

'If ModInit.LocationCode = "001" Then
'    cmdMap.Visible = True
'Else
'    cmdMap.Visible = False
'End If
End Sub

Private Sub txtChqBarCode_Change()
On Error Resume Next
Me.txtCCChequeNo.Text = ""
Me.txtCCChequeAmt.Text = ""
Me.dtCentreDate.Value = ""
Me.dtChqDate.Value = ""
Me.cboChqType.Text = ""
Me.txtCMSFlag.Text = ""
Me.txtCMSNo.Text = ""
Me.txtCMSDate.Text = ""
Me.txtStudentName.Text = ""
txtCentreCode.Text = ""
lblDupliChq.Visible = False
Me.cmsAccept.Enabled = False
cmdShowDupli.Visible = False
End Sub

Private Sub txtChqBarCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cmdSave_Click
    KeyAscii = 0
End If
End Sub
