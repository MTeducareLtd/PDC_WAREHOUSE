VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmActivatePDC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activate PDC"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CheckBox chkStart 
      Caption         =   "Start PDC Factory on"
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
      TabIndex        =   2
      Top             =   1320
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.ComboBox cboCentreCode 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cboInstCode 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComCtl2.DTPicker dtStart 
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
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
      Format          =   138280963
      CurrentDate     =   39310
   End
   Begin VB.ComboBox cboDivision 
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
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin VB.ComboBox cboCentre 
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
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Centre Name"
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
      TabIndex        =   7
      Top             =   720
      Width           =   1110
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Division Name"
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
      TabIndex        =   6
      Top             =   240
      Width           =   1230
   End
End
Attribute VB_Name = "FrmActivatePDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboCentre_Click()
On Error Resume Next
cboCentreCode.ListIndex = cboCentre.ListIndex
End Sub

Private Sub cboDivision_Click()
On Error Resume Next
Me.cboInstCode.ListIndex = cboDivision.ListIndex
FillCentre
End Sub

Private Sub chkStart_Click()
On Error Resume Next
If chkStart.Value = vbChecked Then
    Me.cmdSave.Enabled = True
Else
    Me.cmdSave.Enabled = False
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub FillCentre()
On Error GoTo ErrPath
Me.cboCentre.Clear
Me.cboCentreCode.Clear

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

rs1.Open "Select * from C008_Centers where left(Source_Center_Code,2) ='" & Me.cboInstCode.Text & "' and (PDCFactoryFlag is Null or PDCFactoryFlag =0) order by Source_Center_Name", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        Me.cboCentre.AddItem rs1!Source_Center_Name
        Me.cboCentreCode.AddItem rs1!Source_Center_Code
        rs1.MoveNext
    Loop
    cboCentre.ListIndex = 0
End If
rs1.Close
cn1.Close
Exit Sub

ErrPath:
MsgBox Err.Description
End Sub

Private Sub FillDivision()
On Error Resume Next
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

rs1.Open "Select * from Division order by DivisionName", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        Me.cboDivision.AddItem rs1!DivisionName
        Me.cboInstCode.AddItem rs1!InstituteCode
        
        rs1.MoveNext
    Loop
    cboDivision.ListIndex = 0
End If
rs1.Close
cn1.Close

End Sub

Private Sub cmdSave_Click()
On Error Resume Next
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

rs1.Open "Update g_centre_mis set PDCFactoryFlag =1, PDCFactoryActiveDate ='" & Format(dtStart.Value, "dd Mmm yyyy") & "' where Institutecode ='" & Me.cboInstCode.Text & "' and liccode ='" & Me.cboCentreCode.Text & "'", cn1, adOpenDynamic, adLockPessimistic
cn1.Close

MsgBox "Start Date successfully set.", vbInformation + vbOKOnly
Me.cboCentre.SetFocus

End Sub

Private Sub Form_Load()
FillDivision
dtStart.Value = Date
End Sub
