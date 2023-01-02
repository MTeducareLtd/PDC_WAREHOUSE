VERSION 5.00
Begin VB.Form FrmInwardEntryNew_Sub2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Inward Entry"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Accept"
      Height          =   375
      Left            =   1470
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
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
      Left            =   2550
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3270
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
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
      Left            =   2550
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
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
      Left            =   2580
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   2415
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
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1875
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
      Left            =   270
      TabIndex        =   6
      Top             =   1200
      Width           =   1245
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
      Left            =   270
      TabIndex        =   5
      Top             =   240
      Width           =   1290
   End
End
Attribute VB_Name = "FrmInwardEntryNew_Sub2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrExit
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

rs1.Open "Select * from ASPDC_DispatchSlipLogNew where (Location_Code ='" & ModInit.LocationCode & "' or Location_Code is Null) and DispatchSlipCode not in (Select DispatchSlipCode from ASPDC_DispatchSlip) and InwardNo ='" & Me.txtInwardNo.Text & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    'Update dispatch slip no
    rs2.Open "Select * from ASPDC_DispatchSlipLogNew where InwardNo ='" & Me.txtInwardNo.Text & "'", cn1, adOpenDynamic, adLockPessimistic
    If Not (rs2.BOF And rs2.EOF) Then
        rs2!DispatchSlipCode = Me.txtSlipNo.Text
        rs2!ChequeCnt = Val(txtChequeCnt.Text)
        rs2.Update
        
        
    End If
    rs2.Close
    
    FrmInwardEntryNew.Grid.TextMatrix(FrmInwardEntryNew.Grid.RowSel, 1) = Me.txtSlipNo.Text
    FrmInwardEntryNew.Grid.TextMatrix(FrmInwardEntryNew.Grid.RowSel, 4) = Me.txtChequeCnt.Text
Else
    MsgBox "Dispatch Slip Entry for this Slip is already done. You can't change Cheque Count for this Slip.", vbCritical + vbOKOnly
End If
rs1.Close
cn1.Close
Unload Me
Exit Sub

ErrExit:
MsgBox Err.Description, vbCritical + vbOKOnly
End Sub
