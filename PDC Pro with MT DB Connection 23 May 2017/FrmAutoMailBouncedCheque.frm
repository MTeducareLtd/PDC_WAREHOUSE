VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmAutoMailBouncedCheque 
   Caption         =   "Auto Mail - Bounced Cheque"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   13305
   Begin VB.CommandButton cmdReadEmailId 
      Caption         =   "Read Centre Email"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   720
      Width           =   1695
   End
   Begin VB.FileListBox FileBouncedChqImage 
      Height          =   2040
      Left            =   240
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdSendMail 
      Caption         =   "&Send Mail"
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdReadChqDetails 
      Caption         =   "Read Cheque &Info"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtPath 
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
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.CommandButton cmdReadImages 
      Caption         =   "&Read Images"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   315
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2730
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   4815
      _Version        =   393216
      Cols            =   10
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
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "<Image Name|<Bar Code|<Centre Name|<Student Name|<Stream Name|<Cheque No|<Cheque Date|<Cheque Amount|<To Email Id|<Mail Status"
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
   Begin MSComDlg.CommonDialog cmd 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid GridEMail 
      Height          =   2730
      Left            =   2520
      TabIndex        =   9
      Top             =   3960
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4815
      _Version        =   393216
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
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "<CentreCode|<Email Id"
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
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Browse Pending Image Folder"
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
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "FrmAutoMailBouncedCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
On Error GoTo ErrPath
cmd.FileName = Me.txtPath.Text
'cmd.Flags =  &H200
cmd.Filter = "Picture Files |*.bmp;*.jpg;*.gif;*.wmf"
cmd.ShowOpen

Dim FileActName As String
FileActName = cmd.FileName

Dim PosN As Integer
Dim PathStr As String
PosN = InStrRev(FileActName, "\")
If PosN > 0 Then
    PathStr = Left$(FileActName, PosN)
Else
    PathStr = ""
End If
Me.txtPath.Text = PathStr


Exit Sub
ErrPath:
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdReadChqDetails_Click()
On Error Resume Next

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

If GridEMail.Rows = 1 Then
    MsgBox "Centre Email Id not found.", vbCritical + vbOKOnly
    Exit Sub
End If

For Cnt = 1 To Grid.Rows - 1
    If Grid.TextMatrix(Cnt, 9) <> "Error" Then
        'Read barcode details
        Err.Clear
        rs1.Open "Select * from ASPDC_DispatchSlipDetails where CCChqIdNo ='" & Grid.TextMatrix(Cnt, 1) & "'", cn1, adOpenDynamic, adLockReadOnly
        If Err.Number = 0 Then
            If Not (rs1.BOF And rs1.EOF) Then
                Grid.TextMatrix(Cnt, 5) = rs1!CCChequeNo
                Grid.TextMatrix(Cnt, 7) = rs1!CCChequeAmt
                Grid.TextMatrix(Cnt, 6) = Format(rs1!CCChequeDate, "dd Mmm yyyy")
               
                rs2.Open "select Source_Division_ShortDesc as divisionname, Source_Center_Name as centrename, Source_Center_Code from C008_Centers g inner join C006_Division  d on left(g.Source_Center_Code,2) = d.source_division_code where g.Source_Center_Code ='" & rs1!CMSCenter_Code & "'", cn1, adOpenDynamic, adLockReadOnly
                'rs2.Open "Select streamname, studentname, centername, divisionname, t.InstituteCode ,t.liccode, t.rollno from Tbl_Mtmis_1 T inner join StudentPayment SP on T.Sbentrycode = sp.sbentrycode where sp.ChequeIdNo ='" & rs1!ChqIdNo & "'", cn1, adOpenDynamic, adLockReadOnly
                If Not (rs2.BOF And rs2.EOF) Then
                    Grid.TextMatrix(Cnt, 2) = rs2!DivisionName & " - " & rs2!centrename
                    Grid.TextMatrix(Cnt, 3) = rs1!StudentName
                    Grid.TextMatrix(Cnt, 4) = ""
                
                
                    For ECnt = 1 To GridEMail.Rows - 1
                        If GridEMail.TextMatrix(ECnt, 0) = rs2!Source_Center_Code Then
                            If GridEMail.TextMatrix(ECnt, 1) <> "" Then
                                Grid.TextMatrix(Cnt, 8) = GridEMail.TextMatrix(ECnt, 1)
                            Else
                                Grid.TextMatrix(Cnt, 8) = "No Email Id"
                                Grid.TextMatrix(Cnt, 9) = "Error"
                            End If
                            Exit For
                        End If
                    Next
                Else
                    Grid.TextMatrix(Cnt, 2) = ""
                    Grid.TextMatrix(Cnt, 3) = ""
                    Grid.TextMatrix(Cnt, 4) = ""
                    Grid.TextMatrix(Cnt, 8) = "No Student Data"
                    Grid.TextMatrix(Cnt, 9) = "Error"
                End If
                rs2.Close
                
            Else
                Grid.TextMatrix(Cnt, 2) = "Invalid Barcode"
                Grid.TextMatrix(Cnt, 9) = "Error"
            End If
        Else
            Exit Sub
        End If
        rs1.Close
    End If
Next
cn1.Close
MsgBox "Cheque Info reading activity finished successfully.", vbInformation + vbOKOnly
End Sub

Private Sub cmdReadEmailId_Click()
On Error Resume Next

Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

'Read Centre Email Id
GridEMail.Rows = 1
rs1.Open "Select CentreEMailId, Source_Center_Code from C008_Centers  where PDCFactoryFlag =1", cn1, adOpenDynamic, adLockReadOnly
If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        GridEMail.Rows = GridEMail.Rows + 1
        GridEMail.TextMatrix(GridEMail.Rows - 1, 0) = rs1!Source_Center_Code
        
        If Left(rs1!Source_Center_Code, 2) = "F0" Or Left(rs1!Source_Center_Code, 2) = "Q0" Then
            GridEMail.TextMatrix(GridEMail.Rows - 1, 1) = "" & rs1!CentreEmailId & "; mehekkhan@mteducare.com"
        Else
            GridEMail.TextMatrix(GridEMail.Rows - 1, 1) = "" & rs1!CentreEmailId
        End If
        rs1.MoveNext
    Loop
End If
rs1.Close
cn1.Close

MsgBox "Centre EMail Id successfully downloaded.", vbInformation + vbOKOnly
End Sub

Private Sub cmdReadImages_Click()
On Error Resume Next

FileBouncedChqImage.Path = txtPath.Text
FileBouncedChqImage.Refresh
Dim Barcode As String
With Grid
    .Rows = 1
    For Cnt = 0 To FileBouncedChqImage.ListCount - 1
        Grid.Rows = Grid.Rows + 1
        Grid.TextMatrix(Grid.Rows - 1, 0) = FileBouncedChqImage.List(Cnt)
        Barcode = Left(Grid.TextMatrix(Grid.Rows - 1, 0), 8)
        If IsNumeric(Replace(Barcode, ".j", "")) = True Then
            Grid.TextMatrix(Grid.Rows - 1, 1) = Replace(Barcode, ".j", "")
        Else
            Grid.TextMatrix(Grid.Rows - 1, 1) = "Invalid Filename"
            Grid.TextMatrix(Grid.Rows - 1, 9) = "Error"
        End If
    Next
End With
MsgBox "Image reading activity finished successfully.", vbInformation + vbOKOnly

End Sub

Private Sub cmdSendMail_Click()
On Error Resume Next
Dim FSO As FileSystemObject
Set FSO = New FileSystemObject

For Cnt = 1 To Grid.Rows - 1
    If Grid.TextMatrix(Cnt, 9) <> "Error" And Grid.TextMatrix(Cnt, 8) <> "" And Grid.TextMatrix(Cnt, 8) <> "Sent" Then
        If SendMail(Cnt) = True Then
            'Change status
            Grid.TextMatrix(Cnt, 9) = "Sent"
            
            DoEvents
            FSO.CopyFile Me.txtPath.Text & "/" & Grid.TextMatrix(Cnt, 0), App.Path & "/SentChequesForMail/" & Grid.TextMatrix(Cnt, 0), True
            
            DoEvents
            'Remove image
            FSO.DeleteFile Me.txtPath.Text & "/" & Grid.TextMatrix(Cnt, 0), True
            DoEvents
            
            'Change flag in ASPDC_BounceChequeEntry table
            ChangeMailSentFlag Grid.TextMatrix(Cnt, 1)
            
            
        Else
            Grid.TextMatrix(Cnt, 9) = "Failed"
            DoEvents
        End If
        
    End If
Next
MsgBox "Mail Sending activity finished successfully.", vbInformation + vbOKOnly
End Sub

Private Sub ChangeMailSentFlag(Barcode As String)
On Error Resume Next
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

rs1.Open "Update ASPDC_BounceChequeEntry set MailSentFlag = 1, MailSentDate = getdate() where CCCHQIdNo ='" & Barcode & "'", cn1, adOpenDynamic, adLockPessimistic
cn1.Close

End Sub

Private Function SendMail(ByVal CurRowNo As Integer) As Boolean
On Error GoTo ErrPath
Dim EMailFileName As String
Dim ToEMailId As String
Dim iMsg
Dim iConf
Dim Flds
Dim schema

EMailFileName = Me.txtPath.Text & "/" & Grid.TextMatrix(CurRowNo, 0)
ToEMailId = Grid.TextMatrix(CurRowNo, 8)

SendMail = True
Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields

' send one copy with Google SMTP server (with autentication)
schema = "http://schemas.microsoft.com/cdo/configuration/"
Flds.Item(schema & "sendusing") = 2
Flds.Item(schema & "smtpserver") = "smtp.gmail.com"
Flds.Item(schema & "smtpserverport") = 465
Flds.Item(schema & "smtpauthenticate") = 1
Flds.Item(schema & "sendusername") = "acountech@gmail.com" ' "mis.mteducare@gmail.com"
Flds.Item(schema & "sendpassword") = "mtEducare@123" ' "MTEducare11"
Flds.Item(schema & "smtpusessl") = 1
Flds.Update

With iMsg
.To = ToEMailId ' "ptinbox@rediff.com; delonxavier@mteducare.com"
.From = "acountech@gmail.com" ' "mis.mteducare@gmail.com"
.cc = "AcountechBC@gmail.com" '(password is MT@Ac123)
.Subject = Grid.TextMatrix(CurRowNo, 1) & " -Bounced Cheque of " & Grid.TextMatrix(CurRowNo, 3)
.HTMLBody = "Dear Sir/ Madam <br><br>Please find attached scan copy of Bounced Cheque for " & Grid.TextMatrix(CurRowNo, 3) & " with following details - <br><br>Centre Name : " & Grid.TextMatrix(CurRowNo, 2) & "<br>Cheque Number : " & Grid.TextMatrix(CurRowNo, 5) & "<br>Cheque Date : " & Grid.TextMatrix(CurRowNo, 6) & "<br>Amount : " & Grid.TextMatrix(CurRowNo, 7) & "<br>Barcode : " & Grid.TextMatrix(CurRowNo, 1) & "<br><br>Regards <br><br>Acountech Solutions<br>"
If Trim$(EMailFileName) <> vbNullString Then
    .AddAttachment (EMailFileName)
End If
Set .Configuration = iConf
.Send
End With

Set iMsg = Nothing
Set iConf = Nothing
Set Flds = Nothing
    
SendMail = True
Exit Function

ErrPath:
SendMail = False

End Function


Private Sub Form_Load()
On Error Resume Next
txtPath.Text = App.Path & "\PendingChequesForMail"
GridEMail.Rows = 1
Grid.Rows = 1
End Sub

Private Sub Form_Resize()
On Error Resume Next
Grid.Width = Me.Width - Grid.Left - 360
Shape1.Width = Grid.Width
SetGridWidth
cmdClose.Left = Grid.Left + Grid.Width - cmdClose.Width
Grid.Height = Me.Height - Grid.Top - 1200
cmdClose.Top = Grid.Top + Grid.Height + 120
End Sub


Private Sub SetGridWidth()
On Error Resume Next
With Grid
    For Cnt = 0 To .Cols - 1
        .ColWidth(Cnt) = (.Width - 330) / (.Cols)
    Next
End With
End Sub
