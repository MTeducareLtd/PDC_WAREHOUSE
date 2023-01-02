Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlClient.SqlDataReader
Imports System.Web.UI.Page
Imports System.IO

Partial Class User_Panel_ConfigCity
    Inherits System.Web.UI.Page
    Dim CS As String = ConfigurationManager.AppSettings("connstring")

    Dim conn As New SqlConnection(CS)
    Dim Cmd As New SqlCommand
    Dim Cmd1 As New SqlCommand
    Dim Cmd2 As New SqlCommand
    Dim Cmd3 As New SqlCommand
    
    Protected Sub btnSearch_Click(sender As Object, e As System.EventArgs) Handles btnSearch.Click
        DivSearch.Visible = True
    End Sub


    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            ddlDivision.Items.Clear()
            ddlCentre.Items.Clear()

            conn.Open()
            Dim estr As String = ""
            estr = "select distinct d.DivisionName, d.Institutecode   from division d inner join aspdc_dispatchslip ad on d.Institutecode = ad.misinstitutecode order by d.divisionname"
            Dim da As New SqlDataAdapter(estr, conn)
            Dim ds As New DataSet
            da.Fill(ds, "DivisionName")
            ddlDivision.DataSource = ds.Tables("DivisionName")
            ddlDivision.DataValueField = "Institutecode"
            ddlDivision.DataTextField = "DivisionName"
            ddlDivision.DataBind()
            conn.Close()
        End If
    End Sub

    Protected Sub ddlDivision_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddlDivision.SelectedIndexChanged
        Dim InstituteCode As String
        InstituteCode = ddlDivision.SelectedItem.Value
        ddlCentre.Items.Clear()

        conn.Open()
        Dim estr As String = ""
        estr = "select distinct d.centrename, d.liccode   from g_centre_mis d inner join aspdc_dispatchslip ad on d.Institutecode = ad.misinstitutecode and d.liccode = ad.liccode where ad.misinstitutecode ='" & institutecode & "' order by centrename"
        Dim da As New SqlDataAdapter(estr, conn)
        Dim ds As New DataSet
        da.Fill(ds, "CentreName")
        ddlCentre.DataSource = ds.Tables("CentreName")
        ddlCentre.DataValueField = "LicCode"
        ddlCentre.DataTextField = "CentreName"
        ddlCentre.DataBind()
        conn.Close()
    End Sub

    Protected Sub btnSearchRecord_Click(sender As Object, e As System.EventArgs) Handles btnSearchRecord.Click
        Dim ds As New DataSet
        Dim SerStr As String
        spinner_preview.Visible = True

        Dim InstituteCode As String
        Try
            InstituteCode = ddlDivision.SelectedItem.Value
        Catch
            InstituteCode = ""
            Exit Sub
        End Try

        Dim CentreCnt As Integer
        Dim CentreCode As String
        CentreCode = ""
        For CentreCnt = 0 To ddlCentre.Items.Count - 1
            If ddlCentre.Items(CentreCnt).Selected = True Then
                CentreCode = CentreCode & "'" & ddlCentre.Items(CentreCnt).Value & "',"
            End If
        Next

        If Right(CentreCode, 1) = "," Then CentreCode = Left(CentreCode, Len(CentreCode) - 1)

        If Len(CentreCode) > 1 Then
            CentreCode = " and studentpayment.liccode in (" & CentreCode & ") "

        End If

        SerStr = ""
        If ddlReportType.SelectedIndex = 1 Then 'Detailed
            SerStr = "select t.centername,  t.streamname , t.sbentrycode, t.studentname,  studentpayment.chkno, convert(varchar(50),studentpayment.chkdate, 106) as ChkDate , studentpayment.amountpaid, convert(varchar(50), studentpayment.paydate, 106) as PayDate  " & _
                     "from studentpayment inner join Tbl_MTMIS_1 t on studentpayment.institutecode = t.institutecode and studentpayment.sbentrycode = t.sbentrycode " & _
                     "where studentpayment.institutecode ='" & InstituteCode & "' " & CentreCode & _
                     "and chkstatus  ='Pending' and (ChequeLocCode not in ('02','03','04','06','09','11') or ChequeLocCode is null) and studentpayment.RecordDelFlag =0 and chkdate >='1 Apr 2013' and t.CurrentStudentlFlag = 1 and t.PendingFlag =0 " & _
                     "order by centername, chkno"
        Else
            SerStr = "select t.centername,  '' as streamname , count(*) as SBentryCode, '' as studentname, '' as chkno, '' as chkdate, sum(studentpayment.amountpaidreal) as AmountPaid " & _
                     "from studentpayment inner join Tbl_MTMIS_1 t on studentpayment.institutecode = t.institutecode and studentpayment.sbentrycode = t.sbentrycode " & _
                     "where studentpayment.institutecode ='" & InstituteCode & "' " & CentreCode & _
                     "and chkstatus  ='Pending' and (ChequeLocCode not in ('02','03','04','06','09','11') or ChequeLocCode is null) and studentpayment.RecordDelFlag =0 and chkdate >='1 Apr 2013' and t.CurrentStudentlFlag = 1 and t.PendingFlag =0 " & _
                     "group by t.centername order by centername"
        End If

        Dim da As New SqlDataAdapter(SerStr, conn)
        Try
            da.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                lblRowCnt.Text = ds.Tables(0).Rows.Count

                If ddlReportType.SelectedIndex = 1 Then
                    dlReport.DataSource = ds
                    dlReport.DataBind()
                    dlReport.Visible = True
                    dlReportSummary.Visible = False
                Else
                    dlReportSummary.DataSource = ds
                    dlReportSummary.DataBind()

                    dlReport.Visible = False
                    dlReportSummary.Visible = True
                End If
            Else
                lblRowCnt.Text = 0
                dlReport.DataSource = Nothing
                dlReport.DataBind()

                dlReportSummary.DataSource = Nothing
                dlReportSummary.DataBind()
            End If
            Searchresult.Visible = True
            table_report_wrapper.Visible = True
            spinner_preview.Visible = False
        Catch
            lblRowCnt.Text = "Error in generating report"
            dlReport.DataSource = Nothing
            dlReport.DataBind()

            dlReportSummary.DataSource = Nothing
            dlReportSummary.DataBind()
            spinner_preview.Visible = False
            Exit Sub
        End Try
    End Sub

    Protected Sub btnExport_Click(sender As Object, e As System.EventArgs) Handles btnExport.Click
        Response.Clear()

        Response.AddHeader("content-disposition", "attachment;filename=ChequesNotDispatched.xls")

        Response.Charset = ""


        Response.ContentType = "application/vnd.xls"

        Dim stringWrite As System.IO.StringWriter = New System.IO.StringWriter()

        Dim htmlWrite As System.Web.UI.HtmlTextWriter = New HtmlTextWriter(stringWrite)

        If ddlReportType.SelectedIndex = 1 Then
            dlReport.RenderControl(htmlWrite)
        Else
            dlReportSummary.RenderControl(htmlWrite)
        End If

        Response.Write(stringWrite.ToString())

        Response.End()
    End Sub
End Class
