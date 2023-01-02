VERSION 5.00
Begin VB.Form FrmCMSSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CMS Settings"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboEMailId 
      Height          =   315
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtEMailId 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2640
      Width           =   4215
   End
   Begin VB.TextBox txtPickupCity 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   2160
      Width           =   4215
   End
   Begin VB.TextBox txtPickupPoint 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox txtCompanyCode 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1200
      Width           =   4215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   1695
   End
   Begin VB.ComboBox cboCentreCode 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cboInstCode 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   495
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
   Begin VB.Label lblEmailId 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Centre EMail ID"
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
      Top             =   2640
      Width           =   1350
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pickup City"
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
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pickup Point"
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
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Company Code"
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
      TabIndex        =   12
      Top             =   1200
      Width           =   1275
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   240
      Width           =   1230
   End
End
Attribute VB_Name = "FrmCMSSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboCentre_Click()
On Error Resume Next
cboCentreCode.ListIndex = cboCentre.ListIndex
Me.cboEMailId.ListIndex = cboCentre.ListIndex
End Sub

Private Sub cboCentre_LostFocus()
On Error Resume Next
Me.txtEMailId.Text = cboEMailId.Text

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


rs2.Open "Select * from c008_centers where Source_center_code ='" & cboCentreCode.Text & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs2.BOF And rs2.EOF) Then
    Me.txtCompanyCode.Text = "" & rs2!CMS_Company_Code
    Me.txtPickupPoint.Text = "" & rs2!CMS_Pick_Up_Point
    Me.txtPickupCity.Text = "" & rs2!CMS_Pickup_City
    Me.txtEMailId.Text = "" & rs2!CentreEmailId
Else
    Me.txtCompanyCode.Text = ""
    Me.txtPickupPoint.Text = ""
    Me.txtPickupCity.Text = ""
    Me.txtEMailId.Text = ""
End If
rs2.Close
cn1.Close

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

rs1.Open "Select * from C008_Centers where source_division_code ='" & Me.cboInstCode.Text & "' order by Source_Center_Name", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        Me.cboCentre.AddItem rs1!Source_Center_Name
        Me.cboCentreCode.AddItem rs1!Source_Center_Code
        Me.cboEMailId.AddItem "" & rs1!CentreEmailId
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

rs1.Open "Select * from C006_Division  order by Source_Division_ShortDesc", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        Me.cboDivision.AddItem rs1!Source_Division_ShortDesc
        Me.cboInstCode.AddItem rs1!source_division_code
        
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

rs1.Open "Update c008_centers set CentreEMailId ='" & Me.txtEMailId.Text & "', CMS_Company_Code ='" & Me.txtCompanyCode.Text & "', CMS_Pick_Up_Point ='" & Me.txtPickupPoint.Text & "', CMS_Pickup_City ='" & Me.txtPickupCity.Text & "' where source_center_code ='" & cboCentreCode.Text & "'", cn1, adOpenDynamic, adLockPessimistic

Me.cboEMailId.List(cboEMailId.ListIndex) = txtEMailId.Text

cn1.Close

MsgBox "CMS Settings successfully saved.", vbInformation + vbOKOnly
Me.cboCentre.SetFocus

End Sub

Private Sub Form_Load()
FillDivision
End Sub
