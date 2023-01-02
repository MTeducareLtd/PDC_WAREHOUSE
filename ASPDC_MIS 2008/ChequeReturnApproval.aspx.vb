Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlClient.SqlDataReader
Imports System.Web.UI.Page
Imports System.IO

Partial Class ChequeReturnApproval
    Inherits System.Web.UI.Page
    Dim CS As String = ConfigurationManager.AppSettings("connstring")

    Dim conn As New SqlConnection(CS)
    Dim Cmd As New SqlCommand
    Dim Cmd1 As New SqlCommand
    Dim Cmd2 As New SqlCommand
    Dim Cmd3 As New SqlCommand

    Protected Sub btnSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        DivSearch.Visible = True
    End Sub


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            ddlDivision.Items.Clear()
            'ddlCentre.Items.Clear()

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

    Protected Sub ddlDivision_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlDivision.SelectedIndexChanged
        'Dim InstituteCode As String
        'InstituteCode = ddlDivision.SelectedItem.Value
        'ddlCentre.Items.Clear()

        'conn.Open()
        'Dim estr As String = ""
        'estr = "select distinct d.centrename, d.liccode   from g_centre_mis d inner join aspdc_dispatchslip ad on d.Institutecode = ad.misinstitutecode and d.liccode = ad.liccode where ad.misinstitutecode ='" & institutecode & "' order by centrename"
        'Dim da As New SqlDataAdapter(estr, conn)
        'Dim ds As New DataSet
        'da.Fill(ds, "CentreName")
        'ddlCentre.DataSource = ds.Tables("CentreName")
        'ddlCentre.DataValueField = "LicCode"
        'ddlCentre.DataTextField = "CentreName"
        'ddlCentre.DataBind()
        'conn.Close()
    End Sub

    Private Sub FillResultGrid()
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


        Dim DateRange As String
        DateRange = id_date_range_picker_1.Value

        Dim da As New SqlDataAdapter("SP_ASPDC_ReturnApproval", conn)
        da.SelectCommand.CommandType = CommandType.StoredProcedure

        da.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.DateTime)).Value = DateValue(Left(DateRange, 10))
        da.SelectCommand.Parameters.Add(New SqlParameter("@InstituteCode", SqlDbType.VarChar)).Value = InstituteCode
        da.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.DateTime)).Value = DateValue(Right(DateRange, 10))
        da.SelectCommand.Parameters.Add(New SqlParameter("@ReportType", SqlDbType.Int)).Value = ddlRequestType.SelectedIndex

        Try
            da.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                lblRowCnt.Text = ds.Tables(0).Rows.Count
                dlReport.DataSource = ds
                dlReport.DataBind()

            Else
                lblRowCnt.Text = 0
                dlReport.DataSource = Nothing
                dlReport.DataBind()

            End If
            Searchresult.Visible = True
            table_report_wrapper.Visible = True
            spinner_preview.Visible = False
        Catch
            lblRowCnt.Text = "Error in generating report"
            dlReport.DataSource = Nothing
            dlReport.DataBind()
            spinner_preview.Visible = False
            Exit Sub
        End Try
    End Sub

    Protected Sub btnSearchRecord_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSearchRecord.Click
        FillResultGrid()
    End Sub

    Protected Sub btnExport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExport.Click
        Response.Clear()

        Response.AddHeader("content-disposition", "attachment;filename=ChequesNotDispatched.xls")

        Response.Charset = ""


        Response.ContentType = "application/vnd.xls"

        Dim stringWrite As System.IO.StringWriter = New System.IO.StringWriter()

        Dim htmlWrite As System.Web.UI.HtmlTextWriter = New HtmlTextWriter(stringWrite)

        dlReport.RenderControl(htmlWrite)

        Response.Write(stringWrite.ToString())

        Response.End()
    End Sub

    Protected Sub dlReport_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataListCommandEventArgs) Handles dlReport.ItemCommand
        If e.CommandName = "Approve" Then
            Dim ReturnRequestCode As String
            ReturnRequestCode = e.CommandArgument
            ChangeRequestStatus(ReturnRequestCode, 1)
        ElseIf e.CommandName = "Reject" Then
            Dim ReturnRequestCode As String
            ReturnRequestCode = e.CommandArgument
            ChangeRequestStatus(ReturnRequestCode, 2)
        End If
    End Sub

    Private Sub ChangeRequestStatus(ByVal ReturnRequestCode As String, ByVal NewStatus As Integer)
        conn.Open()

        Dim da1 As New SqlCommand("SP_ASPDC_ReturnApproval_Edit", conn)
        da1.CommandType = CommandType.StoredProcedure
        da1.Parameters.Add(New SqlParameter("@RequestCode", SqlDbType.VarChar, 50)).Value = ReturnRequestCode
        da1.Parameters.Add(New SqlParameter("@ApproveStatus", SqlDbType.Int)).Value = NewStatus
        da1.Parameters.Add(New SqlParameter("@RequestApproveBy", SqlDbType.VarChar, 50)).Value = Request.Cookies("UserName").Value
        da1.ExecuteNonQuery()

        FillResultGrid()
    End Sub
End Class
