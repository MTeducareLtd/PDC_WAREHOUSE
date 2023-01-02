VERSION 5.00
Begin VB.Form FrmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acountech Login - Ver 3.0"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   6705
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboLocationMICRCode 
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
      ItemData        =   "FrmLogin.frx":0000
      Left            =   4200
      List            =   "FrmLogin.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cboLocCode 
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
      ItemData        =   "FrmLogin.frx":0074
      Left            =   4800
      List            =   "FrmLogin.frx":008A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cboLocation 
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
      ItemData        =   "FrmLogin.frx":00E8
      Left            =   3840
      List            =   "FrmLogin.frx":00EA
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtLogin 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      MaxLength       =   100
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox txtPass1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3840
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2640
      TabIndex        =   7
      Top             =   1320
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2640
      TabIndex        =   6
      Top             =   360
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2640
      TabIndex        =   5
      Top             =   840
      Width           =   810
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboLocation_Click()
On Error Resume Next
cboLocCode.ListIndex = cboLocation.ListIndex
Me.cboLocationMICRCode.ListIndex = cboLocation.ListIndex

End Sub

Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdLogin_Click()
On Error GoTo ErrExit
If Me.cboLocation.ListCount = 0 Then
    MsgBox "Select Location.", vbCritical + vbOKOnly
    cboLocation.SetFocus
    Exit Sub
End If

ModInit.Init

Dim cn1 As ADODB.Connection
Dim rs1 As ADODB.Recordset

Set cn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

'UserType 1 = Manager
'UserType 2 = Cheque Entry
'UserType 3 = Slip Entry
'UserType 4 = CMS

Dim str1 As String
str1 = "Select * from ASPDC_UserMaster ASU inner join ASPDC_Location_User ASL on ASU.PDCUserId = ASL.PDCUserId where PDCUserName ='" & Me.txtLogin.Text & "' and PDCPassword ='" & Me.txtPass1.Text & "' and ActiveStatus = 1 and ASL.Location_Code ='" & Me.cboLocCode.Text & "'"
rs1.Open str1, cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    'Check if this user has access to the selected centre

    ModInit.PDCUserName = rs1!PDCUserName
    ModInit.PDCUserType = rs1!PDCUserType
    ModInit.LocationCode = Me.cboLocCode.Text
    ModInit.LocationName = Me.cboLocation.Text
    ModInit.MICRLocationCode = Me.cboLocationMICRCode.Text
Else
    MsgBox "Invalid User Name or Password", vbCritical + vbOKOnly
    txtLogin.Text = ""
    txtPass1.Text = ""
    txtLogin.SetFocus
    rs1.Close
    cn1.Close
    Exit Sub
End If
rs1.Close
cn1.Close

If ModInit.MICRLocationCode = "" Then
    MsgBox "MICR Code for your location is not defined in Configuration data.", vbCritical + vbOKOnly
    End
End If

FrmMain.Show
Unload Me
Exit Sub

ErrExit:
MsgBox "Error : " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrExit
ModInit.Init
ReadLocation

Exit Sub

ErrExit:
MsgBox "Error : " & Err.Description
End Sub

Private Sub ReadLocation()
On Error GoTo ErrExit
Me.cboLocation.Clear
Me.cboLocCode.Clear
cboLocationMICRCode.Clear

Dim cn1 As ADODB.Connection
Dim rs1 As ADODB.Recordset

Set cn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

str1 = "Select * from ASPDC_Location_Master order by Location_Name"
rs1.Open str1, cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        Me.cboLocation.AddItem rs1!Location_Name
        Me.cboLocCode.AddItem rs1!Location_Code
        cboLocationMICRCode.AddItem "" & rs1!MICRLocationCode
        rs1.MoveNext
    Loop
    cboLocation.ListIndex = 0
End If
rs1.Close
cn1.Close

Exit Sub

ErrExit:
MsgBox "Error : " & Err.Description
End Sub

Private Sub txtLogin_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    txtPass1.SetFocus
    KeyAscii = 0
End If
End Sub

Private Sub txtPass1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cmdLogin.SetFocus
    cmdLogin_Click
    KeyAscii = 0
End If
End Sub
