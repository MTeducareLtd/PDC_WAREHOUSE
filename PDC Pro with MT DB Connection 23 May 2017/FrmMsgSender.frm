VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FrmMsgSender 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3105
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4740
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox txtEntryId 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox txtMsg 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtMNo 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin SHDocVwCtl.WebBrowser wbSend 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   4683
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
      Location        =   "http:///"
   End
End
Attribute VB_Name = "FrmMsgSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StartSend As Boolean
Dim ErrCount As Long
Dim CCChqIdNo As String

Public Function SendSMS(MNo As String, Msg As String, Barcode As String)
On Error Resume Next
CCChqIdNo = Barcode


    txtMNo.Text = Replace(MNo, " ", "")
    
    If Len(txtMNo.Text) = 10 Then txtMNo.Text = "91" & txtMNo.Text
    
    txtMsg.Text = Msg
    
    StartSend = True
    wbSend.Navigate2 "http://api.smscountry.com/SMSCwebservice_bulk.aspx?User=finance_tr&passwd=mtel@4321&mobilenumber=" & txtMNo.Text & "&message=" & Msg & "&sid=MTEDU&mtype=N&DR=Y"

    Do While StartSend = True
        DoEvents
    Loop
    DoEvents


'If OutBoxFlag = False Then
'    If ErrCount = 0 Then
'        MsgBox "Message sending finished successfully.", vbInformation + vbOKOnly
'    Else
'        MsgBox "Message sending finished with " & ErrCount & " error(s).", vbCritical + vbOKOnly
'    End If
'    FrmEditor.cmdSend.Enabled = True
'End If
ModInit.WaitFlag = False
Unload FrmMsgSender
End Function

Private Sub Form_Load()
StartSend = False
End Sub

Private Sub wbSend_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
Dim OutputStr As String

OutputStr = "SMS message(s) sent"
OutputStr1 = "OK"

Dim Result As String
If StartSend = True Then
    Result = wbSend.Document.Body.innerText
    
    If InStr(1, Result, OutputStr) > 0 Or InStr(1, Result, OutputStr1) > 0 Then
        'Store message in sent msg list
        UpdateOutput "1", CCChqIdNo
    Else
        UpdateOutput "0", CCChqIdNo
    End If
End If
StartSend = False

End Sub


Private Sub UpdateOutput(SuccessFlag As Integer, Barcode As String)
On Error Resume Next
Dim cn1 As ADODB.Connection
Set cn1 = New ADODB.Connection

cn1.ConnectionString = ModInit.ConnectStringOnline
cn1.Open

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

rs1.Open "Select * from ASPDC_DispatchSlipDetails where CCChqIdNo ='" & Barcode & "'", cn1, adOpenDynamic, adLockPessimistic
If Not (rs1.BOF And rs1.EOF) Then
    If SuccessFlag = 1 Then
        rs1!SMS_Flag = 1
    Else
        rs1!SMS_Flag = 2    'Error in sending sms
    End If
    rs1!SMS_MobileNo = txtMNo.Text
    rs1!SMS_Date = Date
    rs1.Update
End If
rs1.Close
cn1.Close

End Sub
