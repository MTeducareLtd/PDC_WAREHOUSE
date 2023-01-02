Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlClient.SqlDataReader
Imports System.Web.UI.Page
Imports System.IO

Partial Class ChequesInventory
    Inherits System.Web.UI.Page
    Dim CS As String = ConfigurationManager.AppSettings("connstring")

    Dim conn As New SqlConnection(CS)
    Dim Cmd As New SqlCommand
    Dim Cmd1 As New SqlCommand
    Dim Cmd2 As New SqlCommand
    Dim Cmd3 As New SqlCommand
    Dim OpenCnt, InwardCnt, CMSCnt, ReturnCnt, FinalCnt As Long
    Dim OpenVal, InwardVal, CMSVal, ReturnVal, FinalVal As Double

    Protected Sub btnSearch_Click(sender As Object, e As System.EventArgs) Handles btnSearch.Click
        DivSearch.Visible = True
    End Sub

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            ddlDivision.Items.Clear()

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

    Protected Sub btnSearchRecord_Click(sender As Object, e As System.EventArgs) Handles btnSearchRecord.Click
        OpenCnt = 0
        InwardCnt = 0
        CMSCnt = 0
        ReturnCnt = 0
        FinalCnt = 0
        OpenVal = 0
        InwardVal = 0
        CMSVal = 0
        ReturnVal = 0
        FinalVal = 0

        Dim ds As New DataSet
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

        Dim da As New SqlDataAdapter("SP_ASPDC_InventoryDivision", conn)
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add(New SqlParameter("@InstituteCode", SqlDbType.VarChar, 10)).Value = InstituteCode
        da.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.DateTime)).Value = DateValue(Left(DateRange, 10))
        da.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.DateTime)).Value = DateValue(Right(DateRange, 10))


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

    Protected Sub dlReport_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataListItemEventArgs) Handles dlReport.ItemDataBound
        
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            OpenCnt = OpenCnt + Convert.ToDecimal(DataBinder.Eval(e.Item.DataItem, "OpeningStock"))
            InwardCnt = InwardCnt + Convert.ToDecimal(DataBinder.Eval(e.Item.DataItem, "InwardStock"))
            CMSCnt = CMSCnt + Convert.ToDecimal(DataBinder.Eval(e.Item.DataItem, "CMSCnt"))
            ReturnCnt = ReturnCnt + Convert.ToDecimal(DataBinder.Eval(e.Item.DataItem, "ReturnCnt"))
            FinalCnt = FinalCnt + Convert.ToDecimal(DataBinder.Eval(e.Item.DataItem, "FinalStockCnt"))
            OpenVal = OpenVal + Convert.ToDecimal(DataBinder.Eval(e.Item.DataItem, "OpeningStockVal"))
            InwardVal = InwardVal + Convert.ToDecimal(DataBinder.Eval(e.Item.DataItem, "InwardStockVal"))
            CMSVal = CMSVal + Convert.ToDecimal(DataBinder.Eval(e.Item.DataItem, "CMSVal"))
            ReturnVal = ReturnVal + Convert.ToDecimal(DataBinder.Eval(e.Item.DataItem, "ReturnVal"))
            FinalVal = FinalVal + Convert.ToDecimal(DataBinder.Eval(e.Item.DataItem, "FinalStockVal"))
        ElseIf e.Item.ItemType = ListItemType.Footer Then
            Dim OpenCntLabel As Label = CType(e.Item.FindControl("lblF2"), Label)
            OpenCntLabel.Text = OpenCnt.ToString("n")

            Dim OpenValLabel As Label = CType(e.Item.FindControl("lblF3"), Label)
            OpenValLabel.Text = OpenVal.ToString("n")

            Dim InCntLabel As Label = CType(e.Item.FindControl("lblF4"), Label)
            InCntLabel.Text = InwardCnt.ToString("n")

            Dim InValLabel As Label = CType(e.Item.FindControl("lblF5"), Label)
            InValLabel.Text = InwardVal.ToString("n")

            Dim CMSCntLabel As Label = CType(e.Item.FindControl("lblF6"), Label)
            CMSCntLabel.Text = CMSCnt.ToString("n")

            Dim CMSValLabel As Label = CType(e.Item.FindControl("lblF7"), Label)
            CMSValLabel.Text = CMSVal.ToString("n")

            Dim RetCntLabel As Label = CType(e.Item.FindControl("lblF8"), Label)
            RetCntLabel.Text = ReturnCnt.ToString("n")

            Dim RetValLabel As Label = CType(e.Item.FindControl("lblF9"), Label)
            RetValLabel.Text = ReturnVal.ToString("n")

            Dim FinalCntLabel As Label = CType(e.Item.FindControl("lblF10"), Label)
            FinalCntLabel.Text = FinalCnt.ToString("n")

            Dim FinalValLabel As Label = CType(e.Item.FindControl("lblF11"), Label)
            FinalValLabel.Text = FinalVal.ToString("n")

        End If
    End Sub

    Protected Sub btnExport_Click(sender As Object, e As System.EventArgs) Handles btnExport.Click
        Response.Clear()

        Response.AddHeader("content-disposition", "attachment;filename=ChequesInventory.xls")

        Response.Charset = ""


        Response.ContentType = "application/vnd.xls"

        Dim stringWrite As System.IO.StringWriter = New System.IO.StringWriter()

        Dim htmlWrite As System.Web.UI.HtmlTextWriter = New HtmlTextWriter(stringWrite)


        dlReport.RenderControl(htmlWrite)

        Response.Write(stringWrite.ToString())

        Response.End()
    End Sub
End Class
