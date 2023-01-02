VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmDispatchSlipChequeEntry_Sub1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dispatch Slip Cheque Entry"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   11940
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTranCode 
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
      Left            =   9120
      MaxLength       =   9
      TabIndex        =   26
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtMICRNo 
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
      Left            =   7200
      MaxLength       =   9
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtBarcodeNo 
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
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
   Begin VB.ComboBox cboChqType 
      Height          =   315
      ItemData        =   "FrmDispatchSlipChequeEntry_Sub1.frx":0000
      Left            =   9840
      List            =   "FrmDispatchSlipChequeEntry_Sub1.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtChequeAmt 
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
      Left            =   5760
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtChequeDate 
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
      Left            =   3480
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtChequeEntry 
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
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
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
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   480
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
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   480
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
      Left            =   240
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7320
      Width           =   1695
   End
   Begin VB.TextBox txtSlipAmt 
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
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7320
      Width           =   2415
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
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Accept"
      Height          =   375
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7320
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
      Left            =   11160
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "ADD"
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5010
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8837
      _Version        =   393216
      Cols            =   9
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
      FormatString    =   "<Cheque No|<Cheque Date|<Amount|<Type|<Map Flag|<Cheque ID No|<Cheque Barcode No|<Cheque MICR No|<Tran Code"
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
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Barcode No"
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
      TabIndex        =   25
      Top             =   1080
      Width           =   1725
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12000
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MICR Number - Tran Code"
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
      Left            =   7200
      TabIndex        =   24
      Top             =   1080
      Width           =   2250
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Amount"
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
      Left            =   5760
      TabIndex        =   23
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Date (ddmmyyyy)"
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
      Left            =   3480
      TabIndex        =   22
      Top             =   1080
      Width           =   2145
   End
   Begin VB.Label Label3 
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
      TabIndex        =   21
      Top             =   7080
      Width           =   720
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2040
      TabIndex        =   20
      Top             =   1080
      Width           =   1395
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
      Left            =   5760
      TabIndex        =   19
      Top             =   240
      Width           =   1245
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Slip Amount"
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
      Left            =   3000
      TabIndex        =   17
      Top             =   7080
      Width           =   1515
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
      TabIndex        =   14
      Top             =   240
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
      Left            =   3000
      TabIndex        =   13
      Top             =   240
      Width           =   1245
   End
End
Attribute VB_Name = "FrmDispatchSlipChequeEntry_Sub1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboChqType_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cmdAdd_Click
    KeyAscii = 0
    Exit Sub
End If
End Sub

Private Sub cmdAdd_Click()
On Error Resume Next
If Grid.Rows > Val(txtChequeCnt.Text) Then
    MsgBox "You can't add a new Cheque Entry in this slip.", vbCritical + vbOKOnly
    Exit Sub
End If

If Trim(txtBarcodeNo.Text) = "" Then
    MsgBox "Enter Barcode Number of the Cheque", vbCritical + vbOKOnly
    txtBarcodeNo.SetFocus
    Exit Sub
End If

'Check if we have added duplicate entry
Dim Cnt, SCnt As Integer
Dim res As Integer
For Cnt = 1 To Grid.Rows - 1
    If Grid.TextMatrix(Cnt, 0) = txtChequeEntry.Text Then
        res = MsgBox("Duplicate Cheque Entry in the slip.  Do you want to proceed?", vbQuestion + vbYesNo)
        If res = 7 Then
            txtChequeEntry.SetFocus
            Exit Sub
        End If
    End If
Next



If Trim(txtChequeEntry.Text) = "" Then
    MsgBox "Enter Cheque Number", vbCritical + vbOKOnly
    txtChequeEntry.SetFocus
    Exit Sub
End If

If Trim(Me.txtChequeDate.Text) = "" Then
    MsgBox "Enter Cheque Date", vbCritical + vbOKOnly
    txtChequeDate.SetFocus
    Exit Sub
End If

If IsDate(DateValue(Left(txtChequeDate.Text, 2) & "-" & MonthName(Mid(txtChequeDate.Text, 3, 2)) & "-" & Right(txtChequeDate.Text, 4))) = False Then
    MsgBox "Invalid Cheque Date", vbCritical + vbOKOnly
    txtChequeDate.SetFocus
    Exit Sub
End If

Dim CCChqDate As Date
CCChqDate = DateValue(Left(txtChequeDate.Text, 2) & "-" & MonthName(Mid(txtChequeDate.Text, 3, 2)) & "-" & Right(txtChequeDate.Text, 4))

If CCChqDate < DateValue("1 Apr 2013") Or ChqDate > DateValue("31 Mar 2017") Then
    MsgBox "Wrong Cheque Date", vbCritical + vbOKOnly
    txtChequeDate.SetFocus
    Exit Sub
End If

If Trim(txtChequeAmt.Text) = "" Then
    MsgBox "Enter Cheque Amount", vbCritical + vbOKOnly
    txtChequeAmt.SetFocus
    Exit Sub
End If

If IsNumeric(txtChequeAmt.Text) = False Then
    MsgBox "Invalid Cheque amount", vbCritical + vbOKOnly
    txtChequeAmt.SetFocus
    Exit Sub
End If

'If Me.cboChqType.ListIndex = 0 Then
'    MsgBox "Select Cheque Type", vbCritical + vbOKOnly
'    cboChqType.SetFocus
'    Exit Sub
'End If

With Grid
    .Rows = .Rows + 1
    .TextMatrix(.Rows - 1, 0) = Me.txtChequeEntry.Text
    .TextMatrix(.Rows - 1, 1) = Me.txtChequeDate.Text
    .TextMatrix(.Rows - 1, 2) = Me.txtChequeAmt.Text
    .TextMatrix(.Rows - 1, 3) = Me.cboChqType.Text
    .TextMatrix(.Rows - 1, 6) = Me.txtBarcodeNo.Text
    .TextMatrix(.Rows - 1, 7) = Me.txtMICRNo.Text
    .TextMatrix(.Rows - 1, 8) = Me.txtTranCode.Text
End With

txtChequeEntry.Text = ""
txtChequeDate.Text = ""
txtChequeAmt.Text = ""
Me.txtBarcodeNo.Text = ""
cboChqType.ListIndex = 0
txtMICRNo.Text = ""
txtTranCode.Text = ""

Dim TotChqAmt As Double
TotChqAmt = 0
For Cnt = 1 To Grid.Rows - 1
    TotChqAmt = TotChqAmt + Val(Grid.TextMatrix(Cnt, 2))
Next
Me.txtSlipAmt.Text = Format(TotChqAmt, "0.00")

txtBarcodeNo.SetFocus
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
Unload Me
End Sub


Private Sub cmdSave_Click()
On Error Resume Next

If checkvalid = False Then Exit Sub

Me.MousePointer = vbHourglass

'Save in Mirror table on local machine
Dim cn2 As ADODB.Connection
Set cn2 = New ADODB.Connection

cn2.ConnectionString = ModInit.ConnectStringMirror
cn2.Open , , "panatech"

Dim rsM1 As ADODB.Recordset
Set rsM1 = New ADODB.Recordset

rsM1.Open "delete from ASPDC_DispatchSlipDetails where DispatchSlipCode ='" & Trim(UCase(Me.txtSlipNo.Text)) & "'", cn2, adOpenDynamic, adLockPessimistic

rsM1.Open "Select * from ASPDC_DispatchSlipDetails where DispatchSlipCode ='" & Trim(UCase(Me.txtSlipNo.Text)) & "'", cn2, adOpenDynamic, adLockPessimistic
For Cnt = 1 To Grid.Rows - 1
    rsM1.AddNew
    rsM1.Fields("EntryNo").Value = Cnt
    rsM1.Fields("DispatchSlipCode").Value = Trim(UCase(Me.txtSlipNo.Text))
    rsM1.Fields("CCChequeNo").Value = Grid.TextMatrix(Cnt, 0)
    rsM1.Fields("CCChequeAmt").Value = Grid.TextMatrix(Cnt, 2)
    rsM1.Fields("CCChequeDate").Value = Grid.TextMatrix(Cnt, 1)
    rsM1.Fields("CCChequeType").Value = Grid.TextMatrix(Cnt, 3)
    rsM1.Fields("CCCHQIdNo").Value = Grid.TextMatrix(Cnt, 6)
    rsM1.Fields("MICRNumber").Value = Grid.TextMatrix(Cnt, 7)
    rsM1.Fields("TranCode").Value = Grid.TextMatrix(Cnt, 8)
    rsM1.Fields("Location_Code").Value = ModInit.LocationCode
    rsM1.Update
    
    DoEvents
Next
rsM1.Close
cn2.Close

On Error GoTo ErrExit
'Start actual saving
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

'Start mapping the two entries
Dim str As String
str = "Select * from ASPDC_DispatchSlipDetails where DispatchSlipCode ='" & Trim(UCase(Me.txtSlipNo.Text)) & "'"

Dim AutoChqCnt As Integer
AutoChqCnt = 0

rs1.Open str, cn1, adOpenDynamic, adLockPessimistic
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        'Check if this cheque is already present in grid
        For Cnt = 1 To Grid.Rows - 1
            If Val(Grid.TextMatrix(Cnt, 0)) = Val(rs1!CenterChequeNo) And Val(Grid.TextMatrix(Cnt, 0)) <> 0 And Val(Grid.TextMatrix(Cnt, 2)) = rs1!CentreChequeAmt And Grid.TextMatrix(Cnt, 4) = "" Then
                Grid.TextMatrix(Cnt, 4) = "Mapped"
                Grid.TextMatrix(Cnt, 5) = rs1!ChqIdNo
                
                If DateDiff("d", DateValue(rs1!CentreChequeDate), DateValue(Left(Grid.TextMatrix(Cnt, 1), 2) & "-" & MonthName(Mid(Grid.TextMatrix(Cnt, 1), 3, 2)) & "-" & Right(Grid.TextMatrix(Cnt, 1), 4))) > 30 Then
                    Grid.Row = Cnt
                    Grid.Col = 0
                    Grid.CellBackColor = vbRed
                End If
                
                AutoChqCnt = AutoChqCnt + 1
                
                rs1!AutoMapFlag = 1
                rs1!CCChequeNo = rs1!CenterChequeNo
                rs1!CCChequeAmt = rs1!CentreChequeAmt
                rs1!CCChequeDate = DateValue(Left(Grid.TextMatrix(Cnt, 1), 2) & "-" & MonthName(Mid(Grid.TextMatrix(Cnt, 1), 3, 2)) & "-" & Right(Grid.TextMatrix(Cnt, 1), 4))
                rs1!CCChequeType = "Non-ICICI" ' Grid.TextMatrix(Cnt, 3)
                rs1!CCChqIdNo = Grid.TextMatrix(Cnt, 6)
                If Len(Grid.TextMatrix(Cnt, 7)) = 9 Then
                    rs1!MICRNumber = Grid.TextMatrix(Cnt, 7)
                End If
                rs1!TranCode = Grid.TextMatrix(Cnt, 8)
                rs1!EffectDownloadFlag = 1  'No need to download the effect at centre as it is automapped
                rs1.Update
                Exit For
            End If
        Next
        
        DoEvents
        
        rs1.MoveNext
    Loop
End If
rs1.Close

'Store open entries in ASPDC_DispatchSlip_OpenItems
str = "Select * from ASPDC_DispatchSlip_OpenItems where DispatchSlipCode ='" & Trim(UCase(Me.txtSlipNo.Text)) & "'"
rs1.Open str, cn1, adOpenDynamic, adLockPessimistic
For Cnt = 1 To Grid.Rows - 1
    If Grid.TextMatrix(Cnt, 4) = "" Then
        rs1.AddNew
        rs1.Fields("DispatchSlipCode").Value = Trim(txtSlipNo.Text)
        rs1!OpenItemEntryNo = Cnt
        rs1!CCChequeNo = Grid.TextMatrix(Cnt, 0)
        rs1!CCChequeAmt = Grid.TextMatrix(Cnt, 2)
        rs1!CCChequeDate = DateValue(Left(Grid.TextMatrix(Cnt, 1), 2) & "-" & MonthName(Mid(Grid.TextMatrix(Cnt, 1), 3, 2)) & "-" & Right(Grid.TextMatrix(Cnt, 1), 4))
        rs1!CCChequeType = Grid.TextMatrix(Cnt, 3)
        rs1!CCChqIdNo = Grid.TextMatrix(Cnt, 6)
        rs1!MICRNumber = Grid.TextMatrix(Cnt, 7)
        rs1!TranCode = Grid.TextMatrix(Cnt, 8)
        rs1!OpenEntryFlag = 1
        rs1!LinkedCHQIdNo = ""
        rs1.Update
    End If
    DoEvents
    
Next
rs1.Close
    
rs1.Open "Select * from ASPDC_DispatchSlip where DispatchSlipCode ='" & txtSlipNo.Text & "'", cn1, adOpenDynamic, adLockPessimistic
If Not (rs1.BOF And rs1.EOF) Then
    rs1.Fields("ChqEntryFlag").Value = 1
    rs1.Fields("ChqEntryUserName").Value = TxtUserName.Text
    rs1.Fields("ChqEntryDate").Value = Date

    rs1.Fields("AuthEntryFlag").Value = 0
    rs1.Fields("AuthEntryUserName").Value = ""
    rs1.Fields("AuthEntryDate").Value = Date
    
    rs1.Fields("AutoMapChqCnt").Value = AutoChqCnt
    rs1.Fields("ManualMapChqCnt").Value = 0
    rs1.Fields("OpenChqCnt").Value = Val(Me.txtChequeCnt.Text) - AutoChqCnt
    rs1.Update
End If
rs1.Close

'Update Status in Order Engine DB
Dim cmd As ADODB.Command
Set cmd = New ADODB.Command
cmd.ActiveConnection = cn1
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "usp_UpdateDispatchSlipStatus"

cmd.Parameters.Append cmd.CreateParameter("DispatchSlipCode", adVarChar, adParamInput, 50, txtSlipNo.Text)
'cmd.Parameters.Append cmd.CreateParameter("PDC_Status_Id", adVarChar, adParamInput, 50, "01")
'cmd.Parameters.Append cmd.CreateParameter("PDC_Reason_ID", adVarChar, adParamInput, 50, "01")
'cmd.Parameters.Append cmd.CreateParameter("result", adInteger, adParamOutput)
cmd.Execute

'res = cmd("result")


cn1.Close




Me.MousePointer = 0

MsgBox "Dispatch Slip - Cheque Entry successfully done.", vbOKOnly + vbInformation
Unload Me
Exit Sub

ErrExit:
cn1.Close
MsgBox Err.Description
Me.MousePointer = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
txtFlag.Text = "ADD"
TxtUserName.Text = ModInit.PDCUserName
With Grid
    '<Deposit Flag|<Instrument No.|<Instrument Date|>Instrument Amount|<Bank Name|<Name of the Student|<Course|<Form No.|<Academic Year|<RcptCode|<SBEntryCode|<ChequeIdNo
    .ColWidth(0) = (.Width - 350) / (.Cols)
    .ColWidth(1) = .ColWidth(0)
    .ColWidth(2) = .ColWidth(0)
    .ColWidth(3) = .ColWidth(0)
    .ColWidth(4) = .ColWidth(0)
    .ColWidth(5) = .ColWidth(0)
    .ColWidth(6) = .ColWidth(0)
    .ColWidth(7) = .ColWidth(0)
    .Rows = 1
End With
Me.cboChqType.ListIndex = 0
End Sub

Public Sub FillGrid(SlipCode As String)
On Error Resume Next
If SlipCode = "" Then Exit Sub

Me.MousePointer = vbHourglass
Grid.Rows = 1

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringMirror
cn1.Open , , "panatech"

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

'Start mapping the two entries
Dim str As String
str = "Select * from ASPDC_DispatchSlipDetails where DispatchSlipCode ='" & Trim(UCase(SlipCode)) & "'"
rs1.Open str, cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        Grid.Rows = Grid.Rows + 1
        Cnt = Grid.Rows - 1
        
        Grid.TextMatrix(Cnt, 0) = rs1!CCChequeNo
        Grid.TextMatrix(Cnt, 1) = rs1!CCChequeDate
        Grid.TextMatrix(Cnt, 2) = Format(rs1!CCChequeAmt, "0.00")
        Grid.TextMatrix(Cnt, 3) = rs1!CCChequeType
        Grid.TextMatrix(Cnt, 6) = rs1!CCChqIdNo
        Grid.TextMatrix(Cnt, 7) = rs1!MICRNumber
        Grid.TextMatrix(Cnt, 8) = rs1!TranCode
        rs1.MoveNext
    Loop
End If
rs1.Close
cn1.Close

Me.MousePointer = 0
End Sub

Private Function checkvalid() As Boolean
On Error Resume Next
If Trim(Me.txtSlipNo.Text) = "" Then
    MsgBox "Enter Dispatch Slip Number.", vbInformation + vbOKOnly
    txtSlipNo.SetFocus
    checkvalid = False
    Exit Function
End If

If Val(Me.txtChequeCnt.Text) <> Grid.Rows - 1 Then
    MsgBox "You have not entered details of all cheques in the slip.", vbCritical + vbOKOnly
    Grid.SetFocus
    checkvalid = False
    Exit Function
End If

checkvalid = True
End Function


Private Sub txtBarcodeNo_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    If Len(txtBarcodeNo.Text) <> 8 And Len(txtBarcodeNo.Text) <> 6 Then
        MsgBox "Invalid Barcode.", vbCritical + vbOKOnly
        Exit Sub
    Else
        txtChequeEntry.SetFocus
    End If
    KeyAscii = 0
End If
End Sub

Private Sub txtChequeAmt_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    Me.txtMICRNo.SetFocus
    KeyAscii = 0
End If
End Sub

Private Sub txtChequeDate_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    If Len(txtChequeDate.Text) <> 8 Then
        MsgBox "Wrong Cheque Date.", vbCritical + vbOKOnly
        Exit Sub
    Else
        txtChequeAmt.SetFocus
    End If
    KeyAscii = 0
End If
End Sub

Private Sub txtChequeEntry_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    txtChequeDate.SetFocus
    KeyAscii = 0
End If
End Sub

Private Sub txtMICRNo_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    txtTranCode.SetFocus
    KeyAscii = 0
End If
End Sub

Private Sub txtTranCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cmdAdd_Click
    KeyAscii = 0
    Exit Sub
End If
End Sub
