Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlClient.SqlDataReader
Imports System.Web.UI.Page
Imports System.IO

Partial Class ChequesRegister
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

        Dim CentreCode As String
        Try
            CentreCode = ddlCentre.SelectedItem.Value
        Catch
            CentreCode = ""
            Exit Sub
        End Try

        SerStr = ""
        Dim DispatchSlipCnt As Integer
        Dim DispatchSlipCode As String
        DispatchSlipCode = ""
        For DispatchSlipCnt = 0 To ddlSlipNo.Items.Count - 1
            If ddlSlipNo.Items(DispatchSlipCnt).Selected = True Then
                DispatchSlipCode = DispatchSlipCode & "('" & ddlSlipNo.Items(DispatchSlipCnt).Text & "'),"
            End If
        Next

        If Right(DispatchSlipCode, 1) = "," Then DispatchSlipCode = Left(DispatchSlipCode, Len(DispatchSlipCode) - 1)

        Dim da As New SqlDataAdapter("SP_ASPDC_ChequeRegister", conn)
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add(New SqlParameter("@InstituteCode", SqlDbType.VarChar)).Value = InstituteCode
        da.SelectCommand.Parameters.Add(New SqlParameter("@LicCode", SqlDbType.VarChar)).Value = CentreCode
        da.SelectCommand.Parameters.Add(New SqlParameter("@DispatchSlipCode", SqlDbType.VarChar)).Value = DispatchSlipCode
        da.SelectCommand.Parameters.Add(New SqlParameter("@ChequeNo", SqlDbType.VarChar)).Value = "%" & txtChequeNo.Text & "%"
        da.SelectCommand.Parameters.Add(New SqlParameter("@SBEntryCode", SqlDbType.VarChar)).Value = "%" & txtSBEntryCode.Text & "%"
        da.SelectCommand.Parameters.Add(New SqlParameter("@CCCHQIdNo", SqlDbType.VarChar)).Value = "%" & txtBarcode.Text & "%"

        Try
            da.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                lblRowCnt.Text = ds.Tables(0).Rows.Count


                dlReport.DataSource = ds
                dlReport.DataBind()
                dlReport.Visible = True
                dlReportSummary.Visible = False

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
        'Response.Clear()

        'Response.AddHeader("content-disposition", "attachment;filename=ChequesNotDispatched.xls")

        'Response.Charset = ""


        'Response.ContentType = "application/vnd.xls"

        'Dim stringWrite As System.IO.StringWriter = New System.IO.StringWriter()

        'Dim htmlWrite As System.Web.UI.HtmlTextWriter = New HtmlTextWriter(stringWrite)

        'If ddlReportType.SelectedIndex = 1 Then
        '    dlReport.RenderControl(htmlWrite)
        'Else
        '    dlReportSummary.RenderControl(htmlWrite)
        'End If

        'Response.Write(stringWrite.ToString())

        'Response.End()
    End Sub

    Protected Sub ddlCentre_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddlCentre.SelectedIndexChanged
        Dim InstituteCode As String
        InstituteCode = ddlDivision.SelectedItem.Value

        Dim CentreCode As String
        CentreCode = ddlCentre.SelectedItem.Value

        ddlSlipNo.Items.Clear()

        conn.Open()
        Dim estr As String = ""
        estr = "select distinct dispatchslipcode, DispatchDate  from ASPDC_DispatchSlip where misinstitutecode ='" & InstituteCode & "' and liccode ='" & CentreCode & "' and SlipStatus =1 order by DispatchDate desc "
        Dim da As New SqlDataAdapter(estr, conn)
        Dim ds As New DataSet
        da.Fill(ds, "dispatchslipcode")
        ddlSlipNo.DataSource = ds.Tables("dispatchslipcode")
        ddlSlipNo.DataValueField = "DispatchDate"
        ddlSlipNo.DataTextField = "dispatchslipcode"
        ddlSlipNo.DataBind()
        conn.Close()
    End Sub
End Class
