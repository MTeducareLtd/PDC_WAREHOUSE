Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlClient.SqlDataReader
Imports System.Web.UI.Page
Imports System.IO

Partial Class CMSProjections
    Inherits System.Web.UI.Page
    Dim CS As String = ConfigurationManager.AppSettings("connstring")

    Dim conn As New SqlConnection(CS)
    Dim Cmd As New SqlCommand
    Dim Cmd1 As New SqlCommand
    Dim Cmd2 As New SqlCommand
    Dim Cmd3 As New SqlCommand
    Dim TotSci, TotCom, TotUVA, TotSTB, TotSTBM, TotCBSE, TotICSE, TotTot As Double

    Protected Sub btnSearch_Click(sender As Object, e As System.EventArgs) Handles btnSearch.Click
        DivSearch.Visible = True
    End Sub

    Protected Sub btnSearchRecord_Click(sender As Object, e As System.EventArgs) Handles btnSearchRecord.Click
        TotSci = 0
        TotCom = 0
        TotUVA = 0
        TotSTB = 0
        TotSTBM = 0
        TotCBSE = 0
        TotICSE = 0
        TotTot = 0

        Dim ds As New DataSet
        spinner_preview.Visible = True

        Dim DateRange As String
        DateRange = id_date_range_picker_1.Value
        Try
            Dim da As New SqlDataAdapter("SP_ASPDC_CMSProjections", conn)
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.DateTime)).Value = DateValue(Left(DateRange, 10))
            da.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.DateTime)).Value = DateValue(Right(DateRange, 10))



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
            TotSci = TotSci + Convert.ToDecimal(DataBinder.Eval(e.Item.DataItem, "CMSValueSci"))
            TotCom = TotCom + Convert.ToDecimal(DataBinder.Eval(e.Item.DataItem, "CMSValueCom"))
            TotCBSE = TotCBSE + Convert.ToDecimal(DataBinder.Eval(e.Item.DataItem, "CMSValueCBSE"))
            TotICSE = TotICSE + Convert.ToDecimal(DataBinder.Eval(e.Item.DataItem, "CMSValueICSE"))
            TotUVA = TotUVA + Convert.ToDecimal(DataBinder.Eval(e.Item.DataItem, "CMSValueUVA"))
            TotSTB = TotSTB + Convert.ToDecimal(DataBinder.Eval(e.Item.DataItem, "CMSValueSTB"))
            TotSTBM = TotSTBM + Convert.ToDecimal(DataBinder.Eval(e.Item.DataItem, "CMSValueSTBM"))
            TotTot = TotTot + Convert.ToDecimal(DataBinder.Eval(e.Item.DataItem, "CMSValuetotal"))
        ElseIf e.Item.ItemType = ListItemType.Footer Then
            Dim OpenCntLabel As Label = CType(e.Item.FindControl("lblF2"), Label)
            OpenCntLabel.Text = TotSci.ToString("n")

            Dim OpenValLabel As Label = CType(e.Item.FindControl("lblF3"), Label)
            OpenValLabel.Text = TotCom.ToString("n")

            Dim InCntLabel As Label = CType(e.Item.FindControl("lblF4"), Label)
            InCntLabel.Text = TotCBSE.ToString("n")

            Dim InValLabel As Label = CType(e.Item.FindControl("lblF5"), Label)
            InValLabel.Text = TotICSE.ToString("n")

            Dim CMSCntLabel As Label = CType(e.Item.FindControl("lblF6"), Label)
            CMSCntLabel.Text = TotUVA.ToString("n")

            Dim CMSValLabel As Label = CType(e.Item.FindControl("lblF7"), Label)
            CMSValLabel.Text = TotSTBM.ToString("n")

            Dim RetCntLabel As Label = CType(e.Item.FindControl("lblF8"), Label)
            RetCntLabel.Text = TotSTB.ToString("n")

            Dim RetValLabel As Label = CType(e.Item.FindControl("lblF9"), Label)
            RetValLabel.Text = TotTot.ToString("n")
        End If
    End Sub

End Class
