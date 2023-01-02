VERSION 5.00
Begin VB.Form FrmInwardEntryNew_Sub1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inward Entry"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInwardNo 
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
      Left            =   2550
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   240
      Width           =   2415
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
      Left            =   2520
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2415
   End
   Begin VB.ComboBox cboStatus 
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
      ItemData        =   "FrmInwardEntryNew_Sub1.frx":0000
      Left            =   2520
      List            =   "FrmInwardEntryNew_Sub1.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txtChequeCnt 
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
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtSlipDate 
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
      Left            =   2520
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   1695
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
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Accept"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox txtFlag 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "ADD"
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Inward Number"
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
      Top             =   240
      Width           =   1290
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Entry By"
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
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Slip Type"
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
      Top             =   2160
      Width           =   810
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Count"
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
      Top             =   1680
      Width           =   1245
   End
   Begin VB.Label Label1 
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
      Left            =   210
      TabIndex        =   6
      Top             =   720
      Width           =   1875
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dispatch Date"
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
      Top             =   1200
      Width           =   1245
   End
End
Attribute VB_Name = "FrmInwardEntryNew_Sub1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Function MonName(MonthVal As Integer) As String
On Error Resume Next
Select Case MonthVal
Case 1:  MonName = "Jan"
Case 2:  MonName = "Feb"
Case 3:  MonName = "Mar"
Case 4:  MonName = "Apr"
Case 5:  MonName = "May"
Case 6:  MonName = "Jun"
Case 7:  MonName = "Jul"
Case 8:  MonName = "Aug"
Case 9:  MonName = "Sep"
Case 10:  MonName = "Oct"
Case 11:  MonName = "Nov"
Case 12:  MonName = "Dec"
End Select
End Function

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cmdSave_Click
    KeyAscii = 0
End If
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
Unload Me
End Sub


Private Sub cmdSave_Click()
On Error Resume Next

If checkvalid = False Then Exit Sub

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim str As String
    
str = "Select * from ASPDC_DispatchSlipLogNew where DispatchSlipCode ='" & Trim(UCase(Me.txtSlipNo.Text)) & "'"
rs1.Open str, cn1, adOpenDynamic, adLockReadOnly
If (rs1.BOF And rs1.EOF) Then
    AddSlip
Else
    MsgBox "Inward Entry for the Dispatch Slip already exists or Centre side entry for this slip doen't exists.", vbInformation + vbOKOnly, "Error"
    txtSlipNo.SetFocus
    rs1.Close
    Exit Sub
End If
rs1.Close
cn1.Close

End Sub

Private Sub AddSlip()
On Error Resume Next
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset


rs1.Open "Select * from ASPDC_DispatchSlipLogNew where DispatchSlipCode ='" & txtSlipNo.Text & "'", cn1, adOpenDynamic, adLockPessimistic
If (rs1.BOF And rs1.EOF) Then
    rs1.AddNew
    rs1!DispatchSlipCode = txtSlipNo.Text
    rs1!MISInstituteCode = ""
    rs1!LicCode = ""
    rs1!DispatchDate = Left(txtSlipDate.Text, 2) & "-" & MonName(Val(Mid(txtSlipDate.Text, 3, 2))) & "-20" & Right(txtSlipDate.Text, 2)
    rs1!InwardNo = Me.txtInwardNo.Text
    rs1.Fields("InwardFlag").Value = 1
    rs1.Fields("InwardDate").Value = Date
    rs1.Fields("InwardUserName").Value = TxtUserName.Text
    rs1.Fields("ChequeCnt").Value = Val(Me.txtChequeCnt.Text)
    rs1.Fields("SlipType").Value = Me.cboStatus.Text
    rs1.Fields("Location_Code").Value = ModInit.LocationCode
    rs1!PickupArea_Code = FrmInwardEntryNew.cboArea.ListIndex
    rs1.Update
End If
rs1.Close

cn1.Close

MsgBox "Inward Entry successfully done.", vbOKOnly + vbInformation
Unload Me
End Sub



Private Sub Form_Load()
On Error Resume Next
txtFlag.Text = "ADD"
TxtUserName.Text = ModInit.PDCUserName
Me.cboStatus.ListIndex = 0
ReadNewInwardNo
End Sub

Private Sub ReadNewInwardNo()
On Error Resume Next
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim NewNo As String

rs1.Open "Select isnull(max(InwardNo),0) as LastNo  from ASPDC_DispatchSlipLogNew where PickupArea_Code =" & FrmInwardEntryNew.cboArea.ListIndex, cn1, adOpenDynamic, adLockPessimistic
If Not (rs1.BOF And rs1.EOF) Then
    If rs1!LastNo = 0 Then
        'Generate new no
        Select Case FrmInwardEntryNew.cboArea.ListIndex
            Case 0: NewNo = 2001
            Case 1: NewNo = 1001
            Case 3: NewNo = 3001
            Case 2: NewNo = 1
            Case 4: NewNo = 4001
        End Select
    Else
        If FrmInwardEntryNew.cboArea.ListIndex = 3 Then
            If rs1!LastNo = 4000 Then
                NewNo = 30001
            Else
                NewNo = rs1!LastNo + 1
            End If
        ElseIf FrmInwardEntryNew.cboArea.ListIndex = 1 Then
            If rs1!LastNo = 2000 Then
                NewNo = 10001
            Else
                NewNo = rs1!LastNo + 1
            End If
        ElseIf FrmInwardEntryNew.cboArea.ListIndex = 2 Then
            If rs1!LastNo = 1000 Then
                NewNo = 50001
            Else
                NewNo = rs1!LastNo + 1
            End If
        ElseIf FrmInwardEntryNew.cboArea.ListIndex = 0 Then
            If rs1!LastNo = 3000 Then
                NewNo = 20001
            Else
                NewNo = rs1!LastNo + 1
            End If
        Else
            NewNo = rs1!LastNo + 1
        End If
        
        
    End If
End If
rs1.Close
cn1.Close

Me.txtInwardNo.Text = NewNo
End Sub

Private Function checkvalid() As Boolean
On Error Resume Next
If Trim(Me.txtSlipNo.Text) = "" Then
    MsgBox "Enter Dispatch Slip Number.", vbInformation + vbOKOnly
    txtSlipNo.SetFocus
    checkvalid = False
    Exit Function
End If

If Me.cboStatus.ListIndex = 0 Then
    MsgBox "You have not selected any Type for this Dispatch Slip.  Select Dispatch Slip Type.", vbCritical + vbOKOnly
    cboStatus.SetFocus
    checkvalid = False
    Exit Function
End If

checkvalid = True
End Function



Private Sub txtChequeCnt_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cboStatus.SetFocus
    KeyAscii = 0
End If
End Sub




Private Sub txtSlipDate_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    If Len(txtSlipDate.Text) <> 8 Then
        MsgBox "Wrong Slip Date", vbCritical + vbOKOnly
        txtSlipDate.Text = ""
        txtSlipDate.SetFocus
        KeyAscii = 0
    Else
        txtChequeCnt.SetFocus
        KeyAscii = 0
    End If
End If
End Sub

Private Sub txtSlipNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(txtSlipNo.Text) = 13 Then
        txtSlipDate.SetFocus
    Else
        MsgBox "Invalid Slip No.", vbCritical + vbOKOnly
        
    End If
    KeyAscii = 0
End If
End Sub

