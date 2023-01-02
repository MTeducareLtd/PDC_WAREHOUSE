VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FrmDispatchSlip_Sub1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dispatch Slip Entry"
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
   Begin VB.CommandButton cmdGetData 
      Caption         =   "Get Cheque Data"
      Height          =   375
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtChequeEntry2 
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
      TabIndex        =   18
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
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
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
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
      Left            =   5760
      TabIndex        =   2
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
      TabIndex        =   1
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
      TabIndex        =   14
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
      TabIndex        =   12
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
      TabIndex        =   0
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
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "ADD"
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5370
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9472
      _Version        =   393216
      Cols            =   11
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
      FormatString    =   "<PDC Entry Id|<Cheque No Id|<Cheque No|<Amount|<Cheque Date|<MICR No|<Name|<EMailId|<SBEntryCode|<Center Code|<Mobile No"
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
   Begin SHDocVwCtl.WebBrowser wbSend 
      Height          =   735
      Left            =   5520
      TabIndex        =   19
      Top             =   7080
      Visible         =   0   'False
      Width           =   1815
      ExtentX         =   3201
      ExtentY         =   1296
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
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
      TabIndex        =   17
      Top             =   7080
      Width           =   720
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
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
      TabIndex        =   16
      Top             =   1080
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
      Left            =   5760
      TabIndex        =   15
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
      TabIndex        =   13
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   240
      Width           =   1245
   End
End
Attribute VB_Name = "FrmDispatchSlip_Sub1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
On Error Resume Next
Dim TotalChqEntry As String
TotalChqEntry = txtChequeEntry.Text & Me.txtChequeEntry2.Text

If Len(TotalChqEntry) <> 37 Then
    MsgBox "Wrong Cheque Entry", vbCritical + vbOKOnly
    txtChequeEntry.SetFocus
    Exit Sub
End If

If Left(txtChequeEntry.Text, 1) <> Left(txtChequeEntry2.Text, 1) Then
    MsgBox "Mismatch Cheque Entry", vbCritical + vbOKOnly
    txtChequeEntry.SetFocus
    Exit Sub
End If

If Grid.Rows > Val(txtChequeCnt.Text) Then
    MsgBox "You can't add a new Cheque Entry in this slip.", vbCritical + vbOKOnly
    Exit Sub
End If

'Check if we have added duplicate entry
Dim Cnt, SCnt As Integer
Dim ChqNoId, ChqNo, ChqAmt, ChqDate As String

ChqNoId = Mid(TotalChqEntry, 2, 15)
ChqNo = Mid(TotalChqEntry, 21, 6)
ChqAmt = Right(TotalChqEntry, 11)
ChqAmt = Left(ChqAmt, 8)
ChqDate = Mid(TotalChqEntry, 17, 3) & Right(TotalChqEntry, 3)

For Cnt = 1 To Grid.Rows - 1
    If Grid.TextMatrix(Cnt, 1) = ChqNoId Then
        MsgBox "Duplicate Cheque Entry in the slip.", vbCritical + vbOKOnly
        txtChequeEntry.SetFocus
        Exit Sub
    End If
Next


With Grid
    .Rows = .Rows + 1
    .TextMatrix(.Rows - 1, 0) = .Rows - 1
    .TextMatrix(.Rows - 1, 1) = ChqNoId
    .TextMatrix(.Rows - 1, 2) = ChqNo
    .TextMatrix(.Rows - 1, 3) = Format(Val(Left(ChqAmt, 6) & "." & Right(ChqAmt, 2)), "0.00")
    .TextMatrix(.Rows - 1, 4) = Left(ChqDate, 2) & "-" & MonName(Val(Mid(ChqDate, 3, 2))) & "-20" & Right(ChqDate, 2)
End With

txtChequeEntry.Text = ""
txtChequeEntry2.Text = ""

Dim TotChqAmt As Double
TotChqAmt = 0
For Cnt = 1 To Grid.Rows - 1
    TotChqAmt = TotChqAmt + Val(Grid.TextMatrix(Cnt, 3))
Next
Me.txtSlipAmt.Text = Format(TotChqAmt, "0.00")

txtChequeEntry.SetFocus
End Sub

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

Private Sub cmdCancel_Click()
On Error Resume Next
Unload Me
End Sub


Private Sub cmdGetData_Click()

If Trim(txtSlipNo.Text) = "" Then Exit Sub

'Read Dispatch Slip data from server
wbSend.Navigate2 "http://oe.mteducare.com/pdc_management/ImportDispatchData_OrderEngine.aspx?DSC=" & Me.txtSlipNo.Text
DoEvents


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
    
str = "Select * from ASPDC_DispatchSlip where DispatchSlipCode ='" & Trim(UCase(Me.txtSlipNo.Text)) & "'"
rs1.Open str, cn1, adOpenDynamic, adLockReadOnly
If rs1.BOF And rs1.EOF Then
    AddSlip
Else
    MsgBox "The Dispatch Slip Number already exists.", vbInformation + vbOKOnly, "Error"
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


rs1.Open "Select * from ASPDC_DispatchSlip where DispatchSlipCode ='" & txtSlipNo.Text & "'", cn1, adOpenDynamic, adLockPessimistic
rs1.AddNew
rs1.Fields("DispatchSlipCode").Value = Trim(txtSlipNo.Text)
If IsNumeric(Left(Trim(txtSlipNo.Text), 1)) = True Then
    rs1.Fields("MISInstituteCode").Value = Left(Trim(txtSlipNo.Text), 3)
    rs1.Fields("LICCode").Value = Mid(Trim(txtSlipNo.Text), 4, 2)
Else
    rs1.Fields("MISInstituteCode").Value = Left(Trim(txtSlipNo.Text), 2)
    rs1.Fields("LICCode").Value = Mid(Trim(txtSlipNo.Text), 3, 2)
    rs1.Fields("new_institutecode").Value = Left(Trim(txtSlipNo.Text), 2)
    rs1.Fields("new_liccode").Value = Mid(Trim(txtSlipNo.Text), 3, 2)
End If

rs1.Fields("DispatchDate").Value = DateValue(Left(txtSlipDate.Text, 2) & "-" & MonthName(Mid(txtSlipDate.Text, 3, 2)) & "-" & Right(txtSlipDate.Text, 4))
rs1.Fields("ChequeCnt").Value = Val(Me.txtChequeCnt.Text)
rs1.Fields("ChequeValue").Value = Val(Replace(Me.txtSlipAmt.Text, ",", ""))
rs1.Fields("SlipStatus").Value = 1      'Accepted
rs1.Fields("SlipEntryUserName").Value = TxtUserName.Text
rs1.Fields("SlipEntryDate").Value = Date

rs1.Fields("ChqEntryFlag").Value = 0
rs1.Fields("ChqEntryUserName").Value = ""
rs1.Fields("ChqEntryDate").Value = ""

rs1.Fields("AuthEntryFlag").Value = 0
rs1.Fields("AuthEntryUserName").Value = ""
rs1.Fields("AuthEntryDate").Value = ""

rs1.Fields("AutoMapChqCnt").Value = 0
rs1.Fields("ManualMapChqCnt").Value = 0
rs1.Fields("OpenChqCnt").Value = Val(Me.txtChequeCnt.Text)

rs1.Fields("CompleteEntryFlag").Value = 0   'Entry is not complete as cheque entry has to be done
rs1.Fields("Location_Code").Value = ModInit.LocationCode
rs1.Update
rs1.Close

Dim CCnt As Integer
CCnt = 0

rs1.Open "Select * from ASPDC_DispatchSlipDetails where DispatchSlipCode ='" & Trim(txtSlipNo.Text) & "'", cn1, adOpenDynamic, adLockPessimistic
For Cnt = 1 To Grid.Rows - 1
    rs1.AddNew
    rs1.Fields("DispatchSlipCode").Value = Trim(txtSlipNo.Text)
    rs1.Fields("DispatchSlipEntryCode").Value = Cnt
    rs1.Fields("CHQIdNo").Value = Grid.TextMatrix(Cnt, 1)
    rs1.Fields("CenterChequeNo").Value = Grid.TextMatrix(Cnt, 2)
    rs1.Fields("CentreChequeAmt").Value = Grid.TextMatrix(Cnt, 3)
    rs1.Fields("CentreChequeDate").Value = DateValue(Grid.TextMatrix(Cnt, 4))
    
    rs1.Fields("CCChequeNo").Value = ""
    rs1.Fields("CCChequeAmt").Value = ""
    rs1.Fields("CCChequeDate").Value = ""
    If Grid.TextMatrix(Cnt, 5) <> "000000000" Then
        rs1.Fields("MICRNumber").Value = Grid.TextMatrix(Cnt, 5)
    Else
        rs1.Fields("MICRNumber").Value = ""
    End If
    rs1.Fields("AutoMapFlag").Value = 0
    rs1.Fields("EffectDownloadFlag").Value = 0
    rs1.Fields("Location_Code").Value = ModInit.LocationCode
    
    rs1.Fields("StudentName").Value = "" & Grid.TextMatrix(Cnt, 6)
    rs1.Fields("StudentEMailId").Value = "" & Grid.TextMatrix(Cnt, 7)
    
    rs1.Fields("CMSSBEntryCode").Value = "" & Grid.TextMatrix(Cnt, 8)
    rs1.Fields("CMSCenter_Code").Value = "" & Grid.TextMatrix(Cnt, 9)
    rs1.Fields("SMS_MobileNo").Value = "" & Grid.TextMatrix(Cnt, 10)
    
    rs1.Update
Next
rs1.Close

rs1.Open "Update ASPDC_DispatchSlip set CompleteEntryFlag = 1 where DispatchSlipCode ='" & Trim(txtSlipNo.Text) & "'", cn1, adOpenDynamic, adLockPessimistic
cn1.Close

MsgBox "Dispatch Slip Entry successfully done.", vbOKOnly + vbInformation
Unload Me
End Sub



Private Sub Form_Load()
On Error Resume Next
txtFlag.Text = "ADD"
TxtUserName.Text = ModInit.PDCUserName
With Grid
    '<Deposit Flag|<Instrument No.|<Instrument Date|>Instrument Amount|<Bank Name|<Name of the Student|<Course|<Form No.|<Academic Year|<RcptCode|<SBEntryCode|<ChequeIdNo
    .ColWidth(0) = (.Width - 350) / (.Cols - 2)
    .ColWidth(1) = .ColWidth(0)
    .ColWidth(2) = .ColWidth(0)
    .ColWidth(3) = .ColWidth(0)
    .ColWidth(4) = .ColWidth(0)
    .ColWidth(5) = .ColWidth(0)
    .ColWidth(6) = 0
    .ColWidth(7) = 0
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    .ColWidth(10) = 0
    .Rows = 1
End With
End Sub


Private Function checkvalid() As Boolean
On Error Resume Next
If Trim(Me.txtSlipNo.Text) = "" Then
    MsgBox "Enter Dispatch Slip Number.", vbInformation + vbOKOnly
    txtSlipNo.SetFocus
    checkvalid = False
    Exit Function
End If

If Val(Me.txtChequeCnt.Text) <= 0 Then
    MsgBox "You have not selected any Cheque for this Dispatch Slip.  Select Cheque and then Save the entry.", vbCritical + vbOKOnly
    Grid.SetFocus
    checkvalid = False
    Exit Function
End If

'Check if all cheques are scanned or not
If (Grid.Rows - 1) <> Val(Me.txtChequeCnt.Text) Then
    MsgBox "YOu have not entered all Cheque details for this Dispatch Slip.", vbCritical + vbOKOnly
    checkvalid = False
    Exit Function
End If

checkvalid = True
End Function

Private Sub txtChequeCnt_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    txtChequeEntry.SetFocus
    KeyAscii = 0
End If
End Sub

Private Sub txtChequeEntry_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    If Len(txtChequeEntry.Text) = 19 Then
        txtChequeEntry2.SetFocus
    ElseIf Len(txtChequeEntry.Text) <> 0 Then
        MsgBox "Wrong entry", vbCritical + vbOKOnly
    End If
    KeyAscii = 0
End If
End Sub

Private Sub txtChequeEntry2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    If Len(txtChequeEntry2.Text) = 18 Then
        cmdAdd_Click
    ElseIf Len(txtChequeEntry2.Text) <> 0 Then
        MsgBox "Wrong entry", vbCritical + vbOKOnly
    End If
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

Private Sub wbSend_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
If txtSlipNo.Text = "" Then Exit Sub

Grid.Rows = 1
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

rs1.Open "Select * from ASPDC_DispatchSlipFromOrderEngine where dispatchslipno ='" & Me.txtSlipNo.Text & "'", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        With Grid
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = .Rows - 1
            .TextMatrix(.Rows - 1, 1) = rs1!ChequeIdNo
            .TextMatrix(.Rows - 1, 2) = rs1!pay_insnum
            .TextMatrix(.Rows - 1, 3) = Format(rs1!Instr_Amt, "0.00")
            .TextMatrix(.Rows - 1, 4) = Format(rs1!Pay_InstrDate, "dd-Mmm-yyyy") '   Left(rs1!Pay_InstrDate, 2) & "-" & MonName(Val(Mid(rs1!Pay_InstrDate, 3, 2))) & "-20" & Right(rs1!Pay_InstrDate, 2)
            .TextMatrix(.Rows - 1, 5) = rs1!MICRNo
            .TextMatrix(.Rows - 1, 6) = rs1!Con_FirstName & " " & rs1!Con_MidName & " " & rs1!Con_LastName
            .TextMatrix(.Rows - 1, 7) = "" & rs1!EMailId
            
            .TextMatrix(.Rows - 1, 8) = "" & rs1!SBEntryCode
            .TextMatrix(.Rows - 1, 9) = "" & rs1!Center_Code
            .TextMatrix(.Rows - 1, 10) = "" & rs1!Handphone1
            
        End With
    
        rs1.MoveNext
    Loop
End If
rs1.Close



txtChequeEntry.Text = ""
txtChequeEntry2.Text = ""

Dim TotChqAmt As Double
TotChqAmt = 0
For Cnt = 1 To Grid.Rows - 1
    TotChqAmt = TotChqAmt + Val(Grid.TextMatrix(Cnt, 3))
Next
Me.txtSlipAmt.Text = Format(TotChqAmt, "0.00")


cn1.Close
End Sub

