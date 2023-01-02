VERSION 5.00
Begin VB.Form FrmCMSDeposit_Sub2 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Cheque Entry"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   Icon            =   "FrmCMSDeposit_Sub2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   315
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtChqNo 
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
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00004080&
      Caption         =   "Cheque Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1395
   End
End
Attribute VB_Name = "FrmCMSDeposit_Sub2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Private Sub cmdCancel_Click()
'On Error Resume Next
'Unload Me
'End Sub
'
'
'Private Sub CmdShow_Click()
'On Error Resume Next
'If Me.txtChqNo.Text = "" Then Exit Sub
'
'FrmCMSDeposit_Sub1.CmsChkFlag = True
'
'If FrmCMSDeposit_Sub1.PrevVal = Trim(txtChqNo.Text) Then
'    FrmCMSDeposit_Sub1.PrevCnt = Val(FrmCMSDeposit_Sub1.PrevCnt)
'Else
'    FrmCMSDeposit_Sub1.PrevCnt = 1
'End If
'
'Dim FoundFlag As Boolean
'FoundFlag = False
'With FrmCMSDeposit_Sub1.Grid
'    For Cnt = FrmCMSDeposit_Sub1.PrevCnt To .Rows - 1
'        If .TextMatrix(Cnt, 1) = txtChqNo.Text Then
'            .Row = Cnt
'            .RowSel = Cnt
'            .Col = 2
'            .ColSel = .Cols - 1
'
'            If .RowIsVisible(Cnt) = False Then
'                .TopRow = Cnt
'            End If
'            FrmCMSDeposit_Sub1.PrevCnt = Val(Cnt) + 1
'            FrmCMSDeposit_Sub1.PrevVal = Trim(txtChqNo.Text)
'            FoundFlag = True
'            Exit For
'            Exit Sub
'        End If
'    Next
'End With
'
'If FoundFlag = False Then
'    MsgBox "Cheque not found.", vbInformation + vbOKOnly
'    txtChqNo.SetFocus
'    Exit Sub
'Else
'    Unload Me
'    FrmCMSDeposit_Sub1.CmsChkFlag = False
'End If
'End Sub
'
'Private Sub Form_Load()
'On Error Resume Next
'ModInit.SetFormColor Me, "Orange"
'End Sub
'
'
'Private Sub txtChqNo_GotFocus()
'On Error Resume Next
'txtChqNo.SelStart = 0
'txtChqNo.SelLength = Len(txtChqNo.Text)
'End Sub
'
'Private Sub txtChqNo_KeyPress(KeyAscii As Integer)
'On Error Resume Next
'If KeyAscii = 13 Then
'    CmdShow_Click
'    KeyAscii = 0
'End If
'End Sub
