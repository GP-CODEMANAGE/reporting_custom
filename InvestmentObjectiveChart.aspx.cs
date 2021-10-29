using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.IO;
using System.Net;
using System.Collections;
using Spire.Xls;
using System.Drawing;
using System.Data.Common;
using System.Xml;
using iTextSharp.text;
using iTextSharp.text.pdf;

public partial class InvestmentObjectiveChart : System.Web.UI.Page
{
    Boolean fbCheckExcel = false;
    public StreamWriter sw = null;
    string strDescription = string.Empty;
    bool bProceed = true;
    public int liPageSize = 29;//30 -- CHANGE THIS VALUE IN THE GENERATEPDF METHOD WHEN CHANGED HERE.
    //public int liPageSize = 27;
    public string lsStringName = "frutigerce-roman";
    public string lsTotalNumberofColumns, lsDistributionName, lsFamiliesName, lsDateName;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            mvShowReport.ActiveViewIndex = 0;
            FillReportflag();
            fillHousehold();
        }
    }
    public string getFinalSp()
    {
        String lsSQL = "";
        string houseHold = "";
        if (ddlHousehold.SelectedValue != "0")
        {
            houseHold = "'" + ddlHousehold.SelectedItem.Text.Replace("'","''") + "'";
        }
        string AsOfDate = txtAsofdate.Text.Trim() == "" ? "null" : "'" + txtAsofdate.Text.Trim() + "'";
        object CashFlag = ddlCash.SelectedValue == "0" ? "null" : ddlCash.SelectedValue;
        object ReportFlag = ddlReportFlg.SelectedValue == "0" ? "null" : ddlReportFlg.SelectedValue;
        object AllocationGroup = ddlAllocationGroup.SelectedValue == "0" ? "null" : "'" + ddlAllocationGroup.SelectedItem.Text.Replace("'", "''") + "'";
        object Report1and2 = ddlReport1and2.SelectedValue == "" ? "null" : ddlReport1and2.SelectedValue;
        object AllAsset = ddlAllAsset.SelectedValue == "0" ? "null" : ddlAllAsset.SelectedValue;
        if (ddlHousehold.SelectedValue != "")
        {
            lsSQL = "exec GreshamPartners_MSCRM.dbo.SP_R_INVESTMENT_OBJECTIVE_CHART_EXCEL_SMA_NEW_BASEDATA  @HouseholdName  = " + houseHold + ", @AsofDate = " + AsOfDate + ", @GreshamAdvisedFlagTxt = 'TIA',@AllocGroupName = " + AllocationGroup + "";
        }
        else
        {
            lsSQL = "exec SP_R_CONSTRUCTIONCHART '" + ddlHousehold.SelectedItem.Text + "'," + AsOfDate + ", null, " + AllocationGroup + "";//"SP_R_Advent_Report_Other";
        }
        return lsSQL;
    }


    public void FillReportflag()
    {
        //ddlReportGroupflag.Items.Clear();
        //ddlReportgroupflag2.Items.Clear();
        //ddlReportGroupflag.Items.Add(new System.Web.UI.WebControls.ListItem("All", "null"));
        //ddlReportgroupflag2.Items.Add(new System.Web.UI.WebControls.ListItem("All", "null"));

        //ddlReportGroupflag.Items.Add(new System.Web.UI.WebControls.ListItem("Yes", "1"));
        //ddlReportgroupflag2.Items.Add(new System.Web.UI.WebControls.ListItem("Yes", "1"));

        //ddlReportGroupflag.Items.Add(new System.Web.UI.WebControls.ListItem("No", "0"));
        //ddlReportgroupflag2.Items.Add(new System.Web.UI.WebControls.ListItem("No", "0"));
        //ddlReportGroupflag.SelectedValue = "1";
        //ddlReportgroupflag2.SelectedValue = "null";

    }
    public void fillHousehold()
    {
        //ddlHousehold.Items.Add(new ListItem("fdf","dfsdf"));
        DB clsDB = new DB();
        DataSet loDataset = clsDB.getDataSet("sp_s_Get_HouseHoldName");
        ddlHousehold.Items.Clear();
        ddlHousehold.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", ""));
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlHousehold.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][1]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
        }

    }
    public void fillContact()
    {
        DB clsDB = new DB();
        ddlReportFlg.Items.Clear();

        DataSet loDataset = clsDB.getDataSet("sp_r_Household_contact_list @Householdname ='" + ddlHousehold.SelectedItem + "'");
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlReportFlg.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][0]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
        }

    }
    public void AllocationGroup()
    {
        DB clsDB = new DB();
        ddlAllocationGroup.Items.Clear();
        drpAllocationGroupTitle.Items.Clear();
        ddlAllocationGroup.Items.Add(new System.Web.UI.WebControls.ListItem("Select", "0"));
        drpAllocationGroupTitle.Items.Add(new System.Web.UI.WebControls.ListItem("Select", "0"));
        DataSet loDataset = clsDB.getDataSet("SP_S_Advent_Allocation_Group  @Householdname ='" + ddlHousehold.SelectedItem.Text.Replace("'","''") + "'");
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlAllocationGroup.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter]["AllocationGroupName"]), Convert.ToString(loDataset.Tables[0].Rows[liCounter]["AllocationGroupName"])));
        }

    }

    protected void ddlHousehold_SelectedIndexChanged(object sender, EventArgs e)
    {
        // fillContact();
        AllocationGroup();
        fillHouseholdTitle();
    }
    protected void drpAllocationGroupTitle_SelectedIndexChanged(object sender, EventArgs e)
    {
        fillGroupAllocationTitle();
    }
    public void fillGroupAllocationTitle()
    {
        DB clsDB = new DB();
        drpAllocationGroupTitle.Items.Clear();
        //drpAllocationGroupTitle.Items.Add(new System.Web.UI.WebControls.ListItem("Select", "0"));
        if (!String.IsNullOrEmpty(ddlAllocationGroup.SelectedValue))
        {
            //drpAllocationGroupTitle.Items.Add(new System.Web.UI.WebControls.ListItem("Select", ""));
            //  Response.Write("SP_S_AllocationGroupTitle  @AllocationGroupName ='" + ddlAllocationGroup.SelectedValue.Replace("'", "''") + "'");
            DataSet loDataset = clsDB.getDataSet("SP_S_AllocationGroupTitle  @AllocationGroupName ='" + ddlAllocationGroup.SelectedValue.Replace("'", "''") + "'");
            for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
            {
                if (Convert.ToString(loDataset.Tables[0].Rows[liCounter]["Column1"]) == "0")
                {
                    drpAllocationGroupTitle.Items.Add(new System.Web.UI.WebControls.ListItem("Select", "0"));
                }
                else
                {
                    drpAllocationGroupTitle.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][0]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
                }
            }
        }

    }
    public void fillHouseholdTitle()
    {
        DB clsDB = new DB();
        drpHouseHoldReportTitle.Items.Clear();
        if (!String.IsNullOrEmpty(ddlHousehold.SelectedValue))
        {
            //drpHouseHoldReportTitle.Items.Add(new System.Web.UI.WebControls.ListItem("Select", ""));
            //    Response.Write(" SP_S_HouseHoldTitle @HouseHoldName ='" + ddlHousehold.SelectedItem.Text.Replace("'","''") + "'");
            DataSet loDataset = clsDB.getDataSet(" SP_S_HouseHoldTitle @HouseHoldName ='" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "'");
            for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
            {
                drpHouseHoldReportTitle.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][0]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
            }
        }

    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        //if (!String.IsNullOrEmpty(txtpriorperiod.Text))
        //{
        //    DateTime loDatetme2 = new DateTime();
        //    loDatetme2 = Convert.ToDateTime(txtpriorperiod.Text);
        //    DateTime loFindEndofday2 = new DateTime(loDatetme2.Year, loDatetme2.Month, 1).AddMonths(1).AddDays(-1);
        //    if (loDatetme2 != loFindEndofday2)
        //    {
        //        lblError.Text = "Please enter valid Prior Period Comparison";
        //        return;
        //    }
        //}


        DateTime loDatetme1 = new DateTime();
        loDatetme1 = Convert.ToDateTime(txtAsofdate.Text);
        DateTime loFindEndofday1 = new DateTime(loDatetme1.Year, loDatetme1.Month, 1).AddMonths(1).AddDays(-1);
        //if (loDatetme1 != loFindEndofday1)
        //{
        //    lblError.Text = "Please enter valid As Of Date";
        //    return;
        //}


        lblError.Text = "";

        if (RadioButton1.Checked)
        {
            fbCheckExcel = false;
            mvShowReport.ActiveViewIndex = 1;
            Generatereport();
        }
        if (RadioButton2.Checked)
        {
            try
            {
                fbCheckExcel = true;

            }
            catch (Exception ex)
            {
                Response.Write(ex.ToString());
                Response.Write(ex.StackTrace);
            }

            //generatesExcelsheets();
        }

        if (rdbtnPDF.Checked)
        {
            generatePDF();
            //generatePDFNew();
            //generatePDF();
        }
    }
    void Page_PreInit(object sender, System.EventArgs args)
    {
        gvReport.SkinID = "gvReportSkin";
    }
    public void Generatereport()
    {
        String lsSQL = getFinalSp();
        DB clsDB = new DB();
        DataSet lodataset;
        lodataset = null;
        //  Response.Write(lsSQL);
        lodataset = clsDB.getDataSet(lsSQL);
        gvReport.AutoGenerateColumns = false;
        gvReport.Controls.Clear();
        gvReport.Columns.Clear();
        gvReport.DataSource = null;
        gvReport.DataBind();
        foreach (DataColumn dc in lodataset.Tables[0].Columns)
        {

            BoundField newboundfiled = new BoundField();
            newboundfiled.DataField = dc.ColumnName;

            newboundfiled.HeaderText = dc.ColumnName;
            newboundfiled.HtmlEncode = false;//added on 02FEB2011
            newboundfiled.HeaderStyle.CssClass = "dummyheader";
            if (dc.ColumnName.Substring(0, 1) != "_")
            {

                gvReport.Columns.Add(newboundfiled);
            }
            //if (dc.DataType.ToString() == "System.Double")
            if (dc.DataType.FullName == "System.Decimal")//System.Decimal
            {
                newboundfiled.HtmlEncode = false;
                newboundfiled.DataFormatString = "{0:#,###0;(#,###0)}";
                newboundfiled.HeaderStyle.HorizontalAlign = HorizontalAlign.Right;
                newboundfiled.ItemStyle.HorizontalAlign = HorizontalAlign.Right;

            }
            else
            {
                newboundfiled.HeaderStyle.HorizontalAlign = HorizontalAlign.Left;
                newboundfiled.ItemStyle.HorizontalAlign = HorizontalAlign.Left;
            }
            if (dc.ColumnName.Contains("% Assets"))
            {
                newboundfiled.HtmlEncode = false;
                newboundfiled.DataFormatString = "{0:#,###0.0;(#,###0.0)}";
                newboundfiled.HeaderStyle.HorizontalAlign = HorizontalAlign.Right;
                newboundfiled.HeaderStyle.Wrap = false;
                //newboundfiled.ItemStyle.CssClass = "ddcblk";
                newboundfiled.ItemStyle.HorizontalAlign = HorizontalAlign.Right;
            }




        }

        //Assing DataSource to GridView.
        for (int i = 0; i < lodataset.Tables[0].Rows.Count; i++)
        {
            if (lodataset.Tables[0].Rows[i]["_Ssi_UnderlineFlg"].ToString() == "True" || lodataset.Tables[0].Rows[i]["_Ssi_SuperBoldFlg"].ToString() == "True")
            {

                if (i != lodataset.Tables[0].Rows.Count - 1)
                {
                    DataRow newCustomersRow = lodataset.Tables[0].NewRow();
                    newCustomersRow[1] = "";
                    lodataset.Tables[0].Rows.InsertAt(newCustomersRow, i + 1);
                }
            }
        }
        gvReport.DataSource = lodataset;
        gvReport.DataBind();




        for (int i = 0; i < lodataset.Tables[0].Rows.Count; i++)
        {
            if (lodataset.Tables[0].Rows[i]["_Ssi_BoldFlg"].ToString() != "True" && lodataset.Tables[0].Rows[i]["_Ssi_SuperBoldFlg"].ToString() != "True")
            {
                gvReport.Rows[i].Cells[0].Text = "&nbsp;&nbsp;&nbsp;&nbsp;" + gvReport.Rows[i].Cells[0].Text;
                gvReport.Rows[i].Cells[0].Wrap = false;
                if (!fbCheckExcel)
                {
                    gvReport.Rows[i].Cells[0].Width = Unit.Percentage(30);
                }
                else
                {
                    gvReport.Rows[i].Cells[0].Width = Unit.Pixel(285);
                }

                for (int liRowcount = 0; liRowcount < gvReport.Columns.Count; liRowcount++)
                {

                    gvReport.Rows[i].Cells[liRowcount].CssClass = "gvReportss";
                }

            }
            else
            {
                gvReport.Rows[i].Cells[0].Font.Bold = true;
                gvReport.Rows[i].Cells[0].Font.Size = FontUnit.Point(9);
                gvReport.Rows[i].Cells[0].BackColor = System.Drawing.Color.FromName("#D8D8D8;");
                gvReport.Rows[i].Cells[0].CssClass = "greyclass";
                //GridView oGridView = (GridView)gvReport;
                //GridViewRow oGridViewRow = new GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Insert);
                //oGridViewRow.BorderWidth = Unit.Pixel(1);
                //TableCell oTableCell = new TableCell();


                //oGridViewRow.Cells.Add(oTableCell);

                //oGridView.Controls[0].Controls.AddAt(i + liBoldCounter-2, oGridViewRow);
                for (int liRowcount = 0; liRowcount < gvReport.Columns.Count && i != 0; liRowcount++)
                {
                    if (gvReport.Rows[i - 1].Cells[liRowcount].CssClass != "")
                    {

                        gvReport.Rows[i - 1].Cells[liRowcount].CssClass = "gvReportssNo";
                    }
                    if (i > 3)
                        gvReport.Rows[i - 2].Cells[liRowcount].CssClass = "gvReportssNo";
                    if (i > 4)
                        gvReport.Rows[i - 3].Cells[liRowcount].CssClass = "gvReportssBlack";

                }


            }


            if (lodataset.Tables[0].Rows[i]["_Ssi_UnderlineFlg"].ToString() == "True")
            {
                // gvReport.Rows[i].Cells[0].CssClass = "dummy";
                for (int liRowcount = 0; liRowcount < gvReport.Columns.Count; liRowcount++)
                {
                    //                    gvReport.Rows[i].Cells[liRowcount].CssClass = "gvReportssBlack";
                    //  if (!String.IsNullOrEmpty(gvReport.Rows[i].Cells[liRowcount].Text.Replace("&nbsp;", "")))
                    {
                        //gvReport.Rows[i].Cells[liRowcount].Font.Overline = true;
                        //gvReport.Rows[i].Cells[liRowcount].Style.Add("border-top", "thin solid #000000");

                    }
                }
                gvReport.Rows[i].Cells[0].Font.Size = FontUnit.Point(9);
            }

            if (lodataset.Tables[0].Rows[i]["_Ssi_SuperBoldFlg"].ToString() == "True")
            {

                // gvReport.Rows[i].CssClass = "BackgroundColor";  //gvReport.Rows[i].BackColor = System.Drawing.Color.FromName("#B7DDE8");

                for (int liRowcount = 0; liRowcount < gvReport.Columns.Count; liRowcount++)
                {

                    // gvReport.Rows[i].Cells[liRowcount].CssClass = "BackgroundColor";
                    gvReport.Rows[i].Cells[liRowcount].Style.Add("background-color", "#B7DDE8");
                    //gvReport.Rows[i].Cells[liRowcount].BackColor = System.Drawing.Color.FromName("#B7DDE8");
                    gvReport.Rows[i].Cells[liRowcount].Font.Bold = true;
                    gvReport.Rows[i].Cells[liRowcount].VerticalAlign = VerticalAlign.Middle;
                    gvReport.Rows[i].Cells[liRowcount].Style.Add("border-top", "15px solid #ffffff");
                    gvReport.Rows[i].Cells[liRowcount].Style.Add("border-bottom", "15px solid #ffffff");
                    gvReport.Rows[i].Cells[liRowcount].BorderColor = System.Drawing.Color.White;

                }

            }
            if (lodataset.Tables[0].Rows[i]["_Ssi_BoldFlg"].ToString() != "True" && lodataset.Tables[0].Rows[i]["_Ssi_UnderlineFlg"].ToString() != "True" && lodataset.Tables[0].Rows[i]["_Ssi_UnderlineFlg"].ToString() != "True" && lodataset.Tables[0].Rows[i]["_Ssi_SuperBoldFlg"].ToString() != "True")
            {
                gvReport.Rows[i].Cells[0].Text = "&nbsp;&nbsp;&nbsp;&nbsp;" + gvReport.Rows[i].Cells[1].Text;
                gvReport.Rows[i].Cells[0].Wrap = true;

            }




        }
        gvReport.Columns[1].Visible = false;
        int lbCheck;
        for (int i = 0; i < gvReport.Rows.Count; i++)
        {
            lbCheck = 0;
            for (int liColumn = 0; liColumn < gvReport.Columns.Count; liColumn++)
            {

                if (String.IsNullOrEmpty(gvReport.Rows[i].Cells[liColumn].Text.Replace("&nbsp;", "")))
                {
                    lbCheck = lbCheck + 1;
                    //gvReport.Rows[i].Cells[liColumn].Style.Add("border-bottom", "thin dotted #000000");
                }
            }
            if (lbCheck == gvReport.Columns.Count - 1)
            {
                for (int liRemoveCounter = gvReport.Rows[i].Cells.Count - 1; liRemoveCounter > 0; liRemoveCounter--)
                {

                    //   Response.Write("loop counter:"+liRemoveCounter + "<br>");
                    //Response.Write("cells count:" + gvReport.Rows[i].Cells.Count + "<br>");
                    try
                    {
                        gvReport.Rows[i].Cells.Remove(gvReport.Rows[i].Cells[liRemoveCounter]);

                    }
                    catch (Exception ex)
                    {
                        Response.Write(ex.ToString() + "<br>Row:" + i + "<br>Ttaol" + gvReport.Columns.Count + "Cell:" + liRemoveCounter);
                    }


                }
                gvReport.Rows[i].Cells[0].ColumnSpan = lbCheck;

            }

        }
        for (int liCounter = 1; liCounter < gvReport.Columns.Count && gvReport.Rows.Count > 3; liCounter++)
        {
            if (!String.IsNullOrEmpty(Convert.ToString(lodataset.Tables[0].Rows[gvReport.Rows.Count - 1][gvReport.HeaderRow.Cells[liCounter].Text])))
            {
                //Response.Write("<br>"+lodataset.Tables[0].Rows[gvReport.Rows.Count - 1][gvReport.HeaderRow.Cells[liCounter].Text]);
                //if (!String.IsNullOrEmpty(txtpriorperiod.Text) && ddlAlignment.SelectedItem.ToString() == "Horizontal")
                //    gvReport.Rows[gvReport.Rows.Count - 1].Cells[liCounter].Text = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(lodataset.Tables[0].Rows[gvReport.Rows.Count - 1][gvReport.HeaderRow.Cells[liCounter].Text]));

            }
            //if (!String.IsNullOrEmpty(txtpriorperiod.Text) && ddlAlignment.SelectedItem.ToString() == "Horizontal")
            //{
            //    for (int liRowcountss = 0; liRowcountss < gvReport.Columns.Count; liRowcountss++)
            //    {
            //        gvReport.Rows[gvReport.Rows.Count - 1].Cells[liRowcountss].BackColor = System.Drawing.Color.FromName("#FFFFFF");
            //        gvReport.Rows[gvReport.Rows.Count - 2].Cells[liRowcountss].BackColor = System.Drawing.Color.FromName("#FFFFFF");
            //        gvReport.Rows[gvReport.Rows.Count - 1].Cells[liRowcountss].CssClass = "whiteclass";
            //        gvReport.Rows[gvReport.Rows.Count - 2].Cells[liRowcountss].CssClass = "gvReportss";
            //    }

            //}
            try
            {

                //if (!String.IsNullOrEmpty(gvReport.Rows[gvReport.Rows.Count - 1].Cells[liCounter].Text.Replace("&nbsp;", "")) && !String.IsNullOrEmpty(txtpriorperiod.Text) && ddlAlignment.SelectedItem.ToString() == "Horizontal")
                //{



                //    if (gvReport.Rows[gvReport.Rows.Count - 1].Cells[liCounter].Text.Contains("("))
                //    {
                //        gvReport.Rows[gvReport.Rows.Count - 1].Cells[liCounter].Text = gvReport.Rows[gvReport.Rows.Count - 1].Cells[liCounter].Text.Replace(")", "%)");

                //    }
                //    else
                //    {
                //        gvReport.Rows[gvReport.Rows.Count - 1].Cells[liCounter].Text = gvReport.Rows[gvReport.Rows.Count - 1].Cells[liCounter].Text + "%";
                //    }

                //}
            }
            catch { }
        }

        gvReport.HeaderRow.BorderWidth = Unit.Pixel(0);



        for (int i = 0; i < gvReport.Rows.Count; i++)
        {

            if (lodataset.Tables[0].Rows[i]["_Ssi_TabFlg"].ToString() == "True")
            {
                gvReport.Rows[i].Cells[0].Text = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + gvReport.Rows[i].Cells[0].Text;
                gvReport.Rows[i].Cells[0].Wrap = true;
                if (!fbCheckExcel)
                {
                    gvReport.Rows[i].Cells[0].Width = Unit.Percentage(30);
                }
                else
                {
                    gvReport.Rows[i].Cells[0].Width = Unit.Pixel(285);
                }
            }
        }


        for (int liRowcount = 0; liRowcount < gvReport.Columns.Count; liRowcount++)
        {
            if (gvReport.Columns[liRowcount].Visible)
            {
                if (liRowcount == 0)
                {
                    // gvReport.Columns[liRowcount].ItemStyle.Width = Unit.Pixel(400);
                    gvReport.Columns[liRowcount].HeaderStyle.Wrap = false;
                    gvReport.Columns[liRowcount].ItemStyle.Wrap = false;
                }
                else
                {
                    // gvReport.Columns[liRowcount].ItemStyle.Width = Unit.Pixel(100);
                    //gvReport.Columns[liRowcount].HeaderStyle.Width = Unit.Pixel(100);

                    //gvReport.Columns[liRowcount].HeaderStyle.Wrap = true;
                    // gvReport.Columns[liRowcount].HeaderStyle.Height = Unit.Pixel(12);
                    gvReport.Columns[liRowcount].ItemStyle.Wrap = false;
                    //Response.Write(gvReport.Columns[liRowcount].HeaderText);

                }
            }
        }

        for (int i = 0; i < gvReport.Rows.Count; i++)
        {
            for (int liRowcount = 0; liRowcount < gvReport.Columns.Count; liRowcount++)
            {
                if (lodataset.Tables[0].Rows[i]["_ssi_greylineflg"].ToString() == "True")
                {
                    gvReport.Rows[i].Cells[liRowcount].Style.Add("color", "#A5A5A5");
                }

                try
                {
                    //if (i == gvReport.Rows.Count - 1 && !String.IsNullOrEmpty(txtpriorperiod.Text))
                    //{
                    //    gvReport.Rows[i].Cells[liRowcount].CssClass = "PercentageDecimal";
                    //}
                }
                catch
                {
                    // Response.Write("<br>Row: "+i+" Column"+liRowcount);
                }

                if (gvReport.Columns[liRowcount].HeaderText.Contains(" Market Value"))
                {
                    //Response.Write("in");
                    gvReport.Columns[liRowcount].HeaderStyle.Wrap = true;
                    gvReport.Columns[liRowcount].HeaderText = gvReport.Columns[liRowcount].HeaderText.Replace("Market Value", "<br>Market Value");
                }
                if (gvReport.Columns[liRowcount].HeaderText.Contains("% Assets"))
                {
                    try
                    {
                        if (lodataset.Tables[0].Rows[i]["_Ssi_SuperBoldFlg"].ToString() == "True")
                        {
                            //Response.Write("<br> " + String.Format("{0:#,####;(#,####)}", gvReport.Rows[i].Cells[liRowcount].Text)) ;
                            gvReport.Rows[i].Cells[liRowcount].Text = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDouble(gvReport.Rows[i].Cells[liRowcount].Text));
                        }
                        if (gvReport.Rows[i].Cells[liRowcount].CssClass == "gvReportss" && lodataset.Tables[0].Rows[i]["_Ssi_SuperBoldFlg"].ToString() != "True")
                        {
                            gvReport.Rows[i].Cells[liRowcount].CssClass = "ddcblk";

                        }
                        if (gvReport.Rows[i].Cells[liRowcount].CssClass == "gvReportssBlack" && lodataset.Tables[0].Rows[i]["_Ssi_SuperBoldFlg"].ToString() != "True")
                        {
                            gvReport.Rows[i].Cells[liRowcount].CssClass = "ddcblkss";

                        }
                        if (gvReport.Rows[i].Cells[liRowcount].CssClass == "gvReportssNo" && lodataset.Tables[0].Rows[i]["_Ssi_SuperBoldFlg"].ToString() != "True")
                        {
                            gvReport.Rows[i].Cells[liRowcount].CssClass = "ddcblksswhite";

                        }


                    }
                    catch
                    {
                        //Response.Write("<br>Row: "+i+" Column"+liRowcount);
                    }
                }
            }

        }
        gvReport.HeaderStyle.Height = Unit.Pixel(25);
    }



    protected void btnBack_Click(object sender, EventArgs e)
    {
        mvShowReport.ActiveViewIndex = 0;
        //txtAsofdate.Text = "";
        //txtpriorperiod.Text = "";


    }


    protected void BtnExport_Click(object sender, EventArgs e)
    {
        fbCheckExcel = true;
        // generatesExcelsheets();



    }

    public void ExportExcel()
    {
        String FileName = "report";

        Response.Clear();

        Response.Write("<html xmlns:o=\"urn:schemasmicrosoft-com:office:office\" xmlns:x=\"urn:schemas-microsoftcom:office:excel\">");

        Response.Write("<head>");
        Response.Write("<!--[if gte mso 9]><xml>");
        Response.Write("<x:ExcelWorkbook>");
        Response.Write("<x:ExcelWorksheets>");
        Response.Write(" <x:ExcelWorksheet>");
        Response.Write(" <x:Name>report</x:Name>");
        Response.Write(" <x:WorksheetOptions>");
        Response.Write("<x:PageSetup><Layout x:Orientation=\"Landscape\"/><x:/PageSetup>");
        Response.Write(" <x:DisplayPageBreak/>");
        Response.Write(" <x:Print>");
        //  Response.Write(" <x:BlackAndWhite/>");
        //Response.Write(" <x:DraftQuality/>");
        Response.Write(" <x:ValidPrinterInfo/>");
        Response.Write(" <x:PaperSizeIndex>5</x:PaperSizeIndex>");
        Response.Write(" <x:Scale>85</x:Scale>");
        Response.Write(" <x:HorizontalResolution>600</x:HorizontalResolution>");
        // Response.Write(" <x:Gridlines/>");
        //Response.Write(" <x:RowColHeadings/>");
        // Response.Write("(<x:RepeatedRows>$1:$6<x:RepeatedCols>");
        // Response.Write("<x:Formula>=report!$6:$6</x:Formula>");
        Response.Write(" </x:Print>");
        Response.Write(" </x:WorksheetOptions>");
        Response.Write(" </x:ExcelWorksheet>");
        Response.Write(" </x:ExcelWorksheets>");
        Response.Write("</x:ExcelWorkbook>");
        Response.Write("<x:ExcelName>");
        Response.Write("<x:Name>Print_Titles</x:Name>");
        Response.Write("<x:SheetIndex>0</x:SheetIndex>");

        Response.Write("<x:Formula>=report!$A:$G,report!$1:$9</x:Formula>");

        Response.Write("</x:ExcelName>");
        Response.Write("</xml><![endif]-->");

        ///       Response.Write("<style>body {font-family:Frutiger 55 Roman;font-size:8pt} .PercentageDecimal{	background-color:#ffffff;mso-number-format:\\#\\,\\#\\#0\\.0%_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0%\\\\\\) ;}  .ddcblk { border-bottom:1pt solid #F2F2F2;mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;}    .whiteclass {	background-color:#ffffff;mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;} .greyclass {	background-color:#D8D8D8;} .BackgroundColor{	background-color:#B7DDE8;}.dummyheader{padding-left:5px; }.dummy{ border-top:1pt solid #000000;}.Title {	font-family:Frutiger 55 Roman; font-size:18px;	font-weight:normal;	text-decoration:none;}.gvReportss {border-bottom:1pt solid #F2F2F2;mso-number-format:\\#\\,\\#\\#0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\\\\\) ;} .gvReportssNo {border-bottom:1pt solid #ffffff;mso-number-format:\\#\\,\\#\\#0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\\\\\) ;} .gvReportssBlack {border-bottom:thin solid #000000;mso-number-format:\\#\\,\\#\\#0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\\\\\) ;} .ddcblkss {border-bottom:thin solid #000000;  mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;} .ddcblksswhite {border-bottom:thin solid #ffffff;  mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;}");

        //      Response.Write("<style>body {font-family:Frutiger 55 Roman;font-size:8pt} .PercentageDecimal{	background-color:#ffffff;mso-number-format:\\#\\,\\#\\#0\\.0%_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0%\\\\\\) ;}  .ddcblk { border-bottom:thin solid #F2F2F2;mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;}    .whiteclass {	background-color:#ffffff;mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;} .greyclass {	background-color:#D8D8D8;} .BackgroundColor{	background-color:#B7DDE8;}.dummyheader{padding-left:5px; }.dummy{ border-top:thin solid #000000;}.Title {	font-family:Frutiger 55 Roman; font-size:18px;	font-weight:normal;	text-decoration:none;}.gvReportss {border-bottom:thin solid #F2F2F2;mso-number-format:\\#\\,\\#\\#0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\\\\\) ;} .gvReportssNo {border-bottom:thin solid #ffffff;mso-number-format:\\#\\,\\#\\#0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\\\\\) ;} .gvReportssBlack {border-bottom:thin solid #000000;mso-number-format:\\#\\,\\#\\#0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\\\\\) ;} .ddcblkss {border-bottom:thin solid #000000;  mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;} .ddcblksswhite {border-bottom:thin solid #ffffff;  mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;}");
        Response.Write("<style> @page  {   margin:.5in .25in .5in .25in; mso-horizontal-page-align:center; mso-header-margin:.25in; mso-footer-margin:.25in; mso-footer-color:red;;mso-footer-data : '&C&\\0022Frutiger 55 Roman\\,Regular\\0022&8 Page &P of &N &R&\\0022Frutiger 55 Roman\\,italic\\0022&8  &KD8D8D8&D, &T'   } body {font-family:Frutiger 55 Roman;font-size:8pt} .PercentageDecimal{	background-color:#ffffff;mso-number-format:\\#\\,\\#\\#0\\.0%_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0%\\\\\\) ;}  .ddcblk { border-bottom:.5pt hairline #F2F2F2;mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;}    .whiteclass {	background-color:#ffffff;mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;} .greyclass {	background-color:#D8D8D8;} .BackgroundColor{	background-color:#B7DDE8;}.dummyheader{padding-left:5px;height:16px }.dummy{ border-top:thin solid #000000;}.Title {	font-family:Frutiger 55 Roman; font-size:18px;	font-weight:normal;	text-decoration:none;}.gvReportss {border-bottom:.5pt hairline #F2F2F2;mso-number-format:\\#\\,\\#\\#0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\\\\\) ;} .gvReportssNo {border-bottom:thin solid #ffffff;mso-number-format:\\#\\,\\#\\#0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\\\\\) ;} .gvReportssBlack {border-bottom:thin solid #000000;mso-number-format:\\#\\,\\#\\#0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\\\\\) ;} .ddcblkss {border-bottom:thin solid #000000;  mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;} .ddcblksswhite {border-bottom:thin solid #ffffff;  mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;}");

        Response.Write(".familyname { font-family:Frutiger 55 Roman;font-size:14pt;font-weight:bold;height:18.0pt; } ");
        Response.Write("ht25px { height:25px; } .assetdistribution { font-family:Frutiger 55 Roman;font-size:12pt; } ");
        Response.Write(".assDate { font-family:Frutiger 55 Roman;font-size:10pt;font-style:italic; } ");


        Response.Write("</style> ");
        Response.Write("</head>");
        Response.Write("<body>");
        Response.Charset = "";
        Response.AddHeader("content-disposition", string.Format("attachment;filename={0}.xls", FileName));
        // Response.AddHeader("content-disposition", string.Format("attachment;filename={0}.xlsx", FileName));

        //Response.Cache.SetCacheability(HttpCacheability.NoCache);
        Response.ContentType = "application/vnd.ms-excel";

        System.IO.StringWriter stringWrite = new System.IO.StringWriter();
        System.Web.UI.HtmlTextWriter htmlWrite = new HtmlTextWriter(stringWrite);
        gvReport.RenderControl(htmlWrite);
        Response.Write(stringWrite.ToString());
        Response.Write("</body>");
        Response.Write("</html>");
        Response.End();

    }

    public void ExportGVtoExcel(GridView gvexcel, string filename)
    {
        HttpResponse response = HttpContext.Current.Response;
        gvexcel.AllowPaging = false;
        gvexcel.AllowSorting = false;

        response.Clear();
        response.Charset = "";
        response.ContentType = "application/vnd.ms-excel";
        response.AddHeader("content-disposition", string.Format("attachment;filename={0}.xlsx", filename));

        using (StringWriter sw = new StringWriter())
        {
            using (HtmlTextWriter htw = new HtmlTextWriter(sw))
            {
                gvexcel.RenderControl(htw);

                response.Write(sw.ToString());
                response.End();
            }
        }
    }
    public override void VerifyRenderingInServerForm(Control control)
    {
        //this method requires for exportGridtoexcel function
    }

    public void grd_clientview_onitemcommand(object sender, GridViewRowEventArgs e)
    {

        if (e.Row.RowType == DataControlRowType.Header && !String.IsNullOrEmpty(txtAsofdate.Text))
        {


            GridView oGridView = (GridView)sender;
            GridViewRow oGridViewRow = new GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert);
            oGridViewRow.BorderWidth = Unit.Pixel(0);

            String lsfamilyName = ddlHousehold.SelectedItem.Text;

            int liCommaCounter = lsfamilyName.IndexOf(",");
            int liSpaceCounter = lsfamilyName.LastIndexOf(" ");
            if (liCommaCounter > 0 && liSpaceCounter > 0)
                lsfamilyName = lsfamilyName.Substring(0, liCommaCounter) + " " + lsfamilyName.Substring(liSpaceCounter);
            else
                lsfamilyName = lsfamilyName;
            if (ddlAllocationGroup.SelectedValue != "")
            {
                lsfamilyName = ddlAllocationGroup.SelectedItem.Text;
            }
            //if (!String.IsNullOrEmpty(drpHouseHoldReportTitle.SelectedValue))
            //    lsfamilyName = drpHouseHoldReportTitle.SelectedValue;

            //if (!String.IsNullOrEmpty(drpAllocationGroupTitle.SelectedValue))
            //    lsfamilyName = drpAllocationGroupTitle.SelectedValue;


            System.Web.UI.WebControls.Table loTable = new System.Web.UI.WebControls.Table();
            loTable.Width = Unit.Percentage(100);
            loTable.HorizontalAlign = HorizontalAlign.Center;
            TableCell oTableCell = new TableCell();

            TableRow loRow = new TableRow();
            TableCell loCell = new TableCell();
            loCell.Text = "";
            loCell.HorizontalAlign = HorizontalAlign.Center;
            loCell.ColumnSpan = gvReport.Columns.Count - 1;
            loRow.Cells.Add(loCell);
            loTable.Rows.Add(loRow);

            loRow = new TableRow();
            loRow.Height = Unit.Pixel(25);
            loCell = new TableCell();
            loCell.Text = lsfamilyName;
            loCell.CssClass = "familyname";
            loCell.Height = Unit.Pixel(25);
            loCell.ColumnSpan = gvReport.Columns.Count - 1;
            loCell.HorizontalAlign = HorizontalAlign.Center;
            loRow.Cells.Add(loCell);
            loTable.Rows.Add(loRow);

            loRow = new TableRow();
            loCell = new TableCell();
            loCell.Text = "ASSET DISTRIBUTION";
            if (ddlAllAsset.SelectedItem.ToString() != "Horizontal")
                loCell.Text = "ASSET DISTRIBUTION COMPARISON";
            loCell.CssClass = "assetdistribution";
            loCell.HorizontalAlign = HorizontalAlign.Center;
            loCell.ColumnSpan = gvReport.Columns.Count - 1;
            loRow.Cells.Add(loCell);
            loTable.Rows.Add(loRow);



            loRow = new TableRow();
            loCell = new TableCell();
            loCell.Text = Convert.ToDateTime(txtAsofdate.Text).ToString("MMMM dd, yyyy") + "<span style=\"color: #ffffff;\">.</span>" + "</span>";
            loCell.HorizontalAlign = HorizontalAlign.Center;
            loCell.CssClass = "assDate";
            loCell.ColumnSpan = gvReport.Columns.Count - 1;
            loRow.Cells.Add(loCell);
            loTable.Rows.Add(loRow);


            loRow = new TableRow();
            loCell = new TableCell();
            loCell.Text = "";
            loCell.ColumnSpan = gvReport.Columns.Count - 1;
            loCell.HorizontalAlign = HorizontalAlign.Center;
            loRow.Cells.Add(loCell);
            loTable.Rows.Add(loRow);




            //loRow = new TableRow();
            //loCell = new TableCell();
            //loCell.Text = "<br><span style=\"font-family:Frutiger 55 Roman;	font-size:12pt;\">" + "<span  style=\"font-family:Frutiger 55 Roman;	font-size:14pt;font-weight:bold;\" >" + lsfamilyName + "</span>";
            //loCell.Text = oTableCell.Text + "<br>ASSET DISTRIBUTION";
            //loCell.Text = oTableCell.Text + "</span><br><span style=\"font-family:Frutiger 55 Roman;	font-size:10pt;font-style:italic;\">" + " " + Convert.ToDateTime(txtAsofdate.Text).ToString("MMMM dd, yyyy") + "<span style=\"color: #ffffff;\">.</span><br> <span style=\"color: #ffffff;\">.</span> " + "</span>";
            //loCell.HorizontalAlign = HorizontalAlign.Center;
            //loRow.Cells.Add(loCell);
            //loTable.Rows.Add(loRow);







            oTableCell.Controls.Add(loTable);
            oTableCell.Height = Unit.Pixel(25);
            oTableCell.ColumnSpan = gvReport.Columns.Count - 1;
            oTableCell.HorizontalAlign = HorizontalAlign.Center;
            oGridViewRow.CssClass = "ht25px";
            oGridViewRow.Cells.Add(oTableCell);
            oGridView.Controls[0].Controls.AddAt(0, oGridViewRow);


        }
    }

    private DataSet AddTotals(DataSet lodataset)
    {
        if (lodataset.Tables[0].Rows.Count > 0)
        {
            DataRow dr = lodataset.Tables[0].NewRow();

            for (int j = 0; j < lodataset.Tables[0].Rows.Count; j++)
            {
                try
                {
                    for (int k = 0; k < lodataset.Tables[0].Columns.Count; k++)
                    {

                        if (lodataset.Tables[0].Columns[k].ColumnName.Contains("Current Portfolio %"))
                        {
                            if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = 0.0M;

                            dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        }

                        if (lodataset.Tables[0].Columns[k].ColumnName.Contains("Current Portfolio Value"))
                        {
                            if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = 0.0M;

                            dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        }


                        if (lodataset.Tables[0].Columns[k].ColumnName.Contains("Suggested Allocation"))
                        {
                            if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = 0.0M;

                            dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        }

                        //if (lodataset.Tables[0].Columns[k].ColumnName.Contains("Tactical Tilt"))
                        //{
                        //    if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                        //        dr[lodataset.Tables[0].Columns[k].ColumnName] = 0;

                        //    dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        //}

                        if (lodataset.Tables[0].Columns[k].ColumnName.Contains("Tactical Target"))
                        {
                            if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = 0;

                            dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        }


                    }
                }
                catch
                {
 
                }
                
            }



            dr["_LineFlg"] = 2;
            dr["Asset Class"] = "TOTAL";
            lodataset.Tables[0].Rows.Add(dr);
            lodataset.AcceptChanges();


        }
        return lodataset;
    }
    public void SetBorder(Cell foCell, bool IsTop, bool IsBottom, bool IsLeft, bool IsRight)
    {
        if (IsTop == true)
        {
            foCell.BorderWidthTop = 1F;
            foCell.BorderColorTop = new iTextSharp.text.Color(System.Drawing.Color.Black);
        }
        if (IsBottom == true)
        {
            foCell.BorderWidthBottom = 1F;
            foCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);
        }
        if (IsLeft == true)
        {
            foCell.BorderWidthLeft = 1F;
            foCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);
        }
        if (IsRight == true)
        {
            foCell.BorderWidthRight = 1F;
            foCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
        }
        //if (TopBottom == true)
        //{
        //    foCell.BorderWidthBottom = 1F;
        //    foCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);

        //    foCell.BorderWidthLeft = 0.2F;
        //    foCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);

        //    foCell.BorderWidthRight = 0.2F;
        //    foCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
        //}
        //else
        //{
        //    foCell.BorderWidthTop= 1F;
        //    foCell.BorderColorTop = new iTextSharp.text.Color(System.Drawing.Color.Black);

        //    foCell.BorderWidthLeft = 0.2F;
        //    foCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);

        //    foCell.BorderWidthRight = 0.2F;
        //    foCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
        //}

    }
    public void generatePDF()
    {
        liPageSize = 29;
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsFooterTxt = String.Empty;
        //String lsSQL = getFinalSp(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2);
        String lsSQL = getFinalSp();
        // Response.Write(lsSQL);
        newdataset = clsDB.getDataSet(lsSQL);

        for (int i = 0; i < newdataset.Tables[0].Rows.Count; i++)
        {
            DataRow newRow = newdataset.Tables[0].NewRow();
            newRow["Asset Class"] = "Directional";// for borderstyle

            DataRow NonDir = newdataset.Tables[0].NewRow();
            NonDir["Asset Class"] = "Non Directional";// for borderstyle

            if (i != 0)
            {
                if (Convert.ToString(newdataset.Tables[0].Rows[i]["IndicatorFlg"]) == "1" && Convert.ToString(newdataset.Tables[0].Rows[i - 1]["IndicatorFlg"]) != "1")
                {

                    newdataset.Tables[0].Rows.InsertAt(newRow, i);
                    newdataset.Tables[0].AcceptChanges();
                    i++;
                }

                if (Convert.ToString(newdataset.Tables[0].Rows[i]["IndicatorFlg"]) == "2" && Convert.ToString(newdataset.Tables[0].Rows[i - 1]["IndicatorFlg"]) != "2")
                {
                    newdataset.Tables[0].Rows.InsertAt(NonDir, i);
                    newdataset.Tables[0].AcceptChanges();
                    i++;
                }
            }
            else
            {
                if (Convert.ToString(newdataset.Tables[0].Rows[i]["IndicatorFlg"]) == "1")
                {

                    newdataset.Tables[0].Rows.InsertAt(newRow, i);
                    newdataset.Tables[0].AcceptChanges();
                    i++;
                }
                if (Convert.ToString(newdataset.Tables[0].Rows[i]["IndicatorFlg"]) == "2" )
                {
                    newdataset.Tables[0].Rows.InsertAt(NonDir, i);
                    newdataset.Tables[0].AcceptChanges();
                    i++;
                }
            }
           

        }

        for (int j = 0; j < newdataset.Tables[0].Rows.Count; j++)
        {
            if (Convert.ToString(newdataset.Tables[0].Rows[j]["IndicatorFlg"]) == "1")
            {
                newdataset.Tables[0].Rows[j].BeginEdit();
                newdataset.Tables[0].Rows[j]["Asset Class"] = "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + Convert.ToString(newdataset.Tables[0].Rows[j]["Asset Class"]);
                newdataset.Tables[0].Rows[j].EndEdit();
                newdataset.Tables[0].AcceptChanges();
            }

            if (Convert.ToString(newdataset.Tables[0].Rows[j]["IndicatorFlg"]) == "2")
            {
                newdataset.Tables[0].Rows[j].BeginEdit();
                newdataset.Tables[0].Rows[j]["Asset Class"] = "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + Convert.ToString(newdataset.Tables[0].Rows[j]["Asset Class"]);
                newdataset.Tables[0].Rows[j].EndEdit();
                newdataset.Tables[0].AcceptChanges();
            }
        }


        newdataset = AddTotals(newdataset);

        if (newdataset.Tables[0].Rows.Count < 1)
        {
            lblError.Text = "No Record Found";
            return;
        }

        DataSet loInsertblankRow = newdataset.Copy();

        newdataset = loInsertblankRow.Clone();

        // string strGUID = Guid.NewGuid().ToString();
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        // strGUID = strGUID.Substring(0, 5);
        // String fsFinalLocation = @"C:\Reports\" + strGUID + ".xls";

        String fsFinalLocation = Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";
        int liBlankCounter = 0;

        for (int liBlankRow = 0; liBlankRow < loInsertblankRow.Tables[0].Columns.Count; liBlankRow++)
        {
            if (liBlankRow == 7 )
            {
                loInsertblankRow.Tables[0].Columns.RemoveAt(liBlankRow);

                for (int i = 0; i < loInsertblankRow.Tables[0].Columns.Count; i++)
                {
                    if (i == 7)
                    {
                        loInsertblankRow.Tables[0].Columns.RemoveAt(i);
                    }
                }
            }


        }

        loInsertblankRow.Tables[0].AcceptChanges();
        DataSet lodataset = new DataSet();
        lodataset = loInsertblankRow.Copy();
        DataSet loInsertdataset = lodataset.Copy();
        for (int liNewdataset = lodataset.Tables[0].Columns.Count - 1; liNewdataset > -1; liNewdataset--)
        {
            if (lodataset.Tables[0].Columns[liNewdataset].ColumnName.Contains("_") || lodataset.Tables[0].Columns[liNewdataset].ColumnName.Trim().Equals("1"))
            {
                loInsertdataset.Tables[0].Columns.Remove(loInsertdataset.Tables[0].Columns[liNewdataset]);
            }
        }
        //    loInsertdataset.Tables[0].Columns.Remove(loInsertdataset.Tables[0].Columns[1]);
        loInsertdataset.AcceptChanges();

        //iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 30, 31, 10);
        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 30, 31, 8);//10,10
        String ls = Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));
        document.Open();


        lsTotalNumberofColumns = loInsertdataset.Tables[0].Columns.Count + "";
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(loInsertdataset.Tables[0].Columns.Count, loInsertdataset.Tables[0].Rows.Count);   // 2 rows, 2 columns           
        iTextSharp.text.Cell loCell = new Cell();
        loTable.Cellpadding = 0f;
        loTable.Cellspacing = 0f;


        setTableProperty(loTable);
        String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();
        int liTotalPage = (loInsertdataset.Tables[0].Rows.Count / liPageSize);
        int liCurrentPage = 0;
        if (loInsertdataset.Tables[0].Rows.Count % liPageSize != 0)
        {
            liTotalPage = liTotalPage + 1;
        }
        else
        {
            liPageSize = 28;
            liTotalPage = liTotalPage + 1;
        }

        //check the length of the column name to set the pagesize.
        for (int j = 0; j < loInsertdataset.Tables[0].Columns.Count; j++)
        {
            if (loInsertdataset.Tables[0].Columns[j].ColumnName.Length > 30)
            {
                liPageSize = 28;
            }
        }

        String SQL = getFinalSp();
        // Response.Write(lsSQL);
        newdataset = clsDB.getDataSet(SQL);

        newdataset = AddTotals(newdataset);

        for (int liRowCount = 0; liRowCount < loInsertblankRow.Tables[0].Rows.Count; liRowCount++)
        {
            if (liRowCount % liPageSize == 0)
            {
                document.Add(loTable);

                if (liRowCount != 0)
                {
                    liCurrentPage = liCurrentPage + 1;
                    document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, liPageSize, false, String.Empty));
                    document.NewPage();
                }

                loInsertdataset.AcceptChanges();
                setHeader(document, loInsertdataset, newdataset);
                loTable = new iTextSharp.text.Table(loInsertdataset.Tables[0].Columns.Count, loInsertdataset.Tables[0].Rows.Count);   // 2 rows, 2 columns           
                setTableProperty(loTable);
            }




            int colsize = loInsertdataset.Tables[0].Columns.Count;
            for (int liColumnCount = 0; liColumnCount < colsize; liColumnCount++)
            {
                iTextSharp.text.Chunk lochunk = new Chunk();
                String lsFormatedString = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]);
                try
                {
                    if (liColumnCount == 1)
                    {
                        lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(lsFormatedString));
                    }
                    else if (lsFormatedString == "0" && liColumnCount == 4)
                    {
                        lsFormatedString = String.Format("", Convert.ToDecimal(lsFormatedString));
                    }
                    //else if (lsFormatedString == "0" && liColumnCount == 3)
                    //{
                    //    lsFormatedString = String.Format("", Convert.ToDecimal(lsFormatedString));
                    //}
                    else if (liColumnCount == loInsertdataset.Tables[0].Columns.Count)
                    {
                        lsFormatedString = String.Format("${0:#,###0.0;(#,###0.0)}%", Convert.ToDecimal(lsFormatedString));
                    }
                    else
                    {
                        lsFormatedString = String.Format("{0:#,###0.0;(#,###0.0)}%", Convert.ToDecimal(lsFormatedString));
                    }
                }
                catch
                {

                }

                //changed on 02/25/2011
                //lochunk = new Chunk(lsFormatedString, Font8Whitecheck(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][0])));
                lochunk = new Chunk(lsFormatedString, Font7Whitecheck(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][0])));
                loCell = new iTextSharp.text.Cell();
                loCell.Border = 0;
                loCell.NoWrap = true;
                //loCell.VerticalAlignment=0;
                loCell.VerticalAlignment = 5;

                setGreyBorder(lodataset, loCell, liRowCount);
                loCell.Leading = 4f;//6

                loCell.UseBorderPadding = true;

                if (liColumnCount == 1 || liColumnCount == 4 || liColumnCount == 3 || liColumnCount == 5 || liColumnCount == 6)
                {
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                }
                else if (liColumnCount == 2)
                {
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                }

                if (liColumnCount == 2 || liColumnCount ==3 || liColumnCount ==5 || liColumnCount == 6)
                {

                    if (Convert.ToString(loInsertblankRow.Tables[0].Rows[liRowCount]["_LineFlg"]) == "1")
                    {
                        //loCell.EnableBorderSide(2);
                    }
                    else if (Convert.ToString(loInsertblankRow.Tables[0].Rows[liRowCount]["_LineFlg"]) == "2")
                    {
                        try
                        {
                            loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                            //loCell.EnableBorderSide(1);
                            if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount]) == "0" && liColumnCount ==5)
                            {
                                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                //loCell.EnableBorderSide(1);
                                string CurrentAllocation = String.Format("", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                lochunk = new Chunk(CurrentAllocation, setFontsAll(8, 1, 0));
                            }
                            else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount]) != "0")
                            {
                                string CurrentAllocation = String.Format("{0:#,###0;(#,###0.0)}%", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                lochunk = new Chunk(CurrentAllocation, setFontsAll(8, 1, 0));
                            }
                            
                        }
                        catch
                        {

                        }
                    }
                }
                else if (liColumnCount == 1 || liColumnCount == 4)
                {
                    if (Convert.ToString(loInsertblankRow.Tables[0].Rows[liRowCount]["_LineFlg"]) == "1")
                    {
                        //loCell.EnableBorderSide(1);
                    }
                    else if (Convert.ToString(loInsertblankRow.Tables[0].Rows[liRowCount]["_LineFlg"]) == "2")
                    {
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                        //loCell.EnableBorderSide(1);
                        if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount]) != "")
                        {
                            string CurrentAllocation = String.Format("${0:#,###0;(#,###0.0)}", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount])));
                            //string SuggestedAllocation = String.Format("{0:#,###0;(#,###0.0)}%", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["Suggested Allocation"])));
                            lochunk = new Chunk(CurrentAllocation, setFontsAll(8, 1, 0));
                            //lochunk = new Chunk(SuggestedAllocation, setFontsAll(8, 1, 0));
                        }
                    }
                }
                else if (liColumnCount == 0)
                {
                    if (Convert.ToString(loInsertblankRow.Tables[0].Rows[liRowCount]["_LineFlg"]) == "2")
                    {
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                        //loCell.EnableBorderSide(1);
                        string CurrentAllocation = Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount]);
                        //string SuggestedAllocation = String.Format("{0:#,###0;(#,###0.0)}%", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["Suggested Allocation"])));
                        lochunk = new Chunk(CurrentAllocation, setFontsAll(8, 1, 0));
                        //lochunk = new Chunk(SuggestedAllocation, setFontsAll(8, 1, 0));
                    }
             
                }

                //else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_LineFlg"]) == "0")
                //{
                //    loCell.DisableBorderSide(-1);
                //}
                /////

                #region Not in Use
                /*=========START WITH BOLD AND SUPERBOLD FLAG========*/
                //if (checkTrue(lodataset, liRowCount, "_OrderNmb") || checkTrue(lodataset, liRowCount, "_OrderNmb"))
                //{
                //    lsFormatedString = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]);
                //    try
                //    {
                //        if (liColumnCount == loInsertdataset.Tables[0].Columns.Count - 1)
                //        {
                //            lsFormatedString = String.Format("${0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(lsFormatedString));
                //        }
                //        else
                //        {
                //            lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(lsFormatedString));
                //        }
                //    }
                //    catch
                //    {

                //    }

                //    //changed on 02/25/2011
                //    //lochunk = new Chunk(lsFormatedString, Font9Bold());
                //    lochunk = new Chunk(lsFormatedString, Font8Bold());
                //    #region Commented
                //    if (!lodataset.Tables[0].Rows[liRowCount][0].ToString().Contains("NET CHANGE"))
                //    {
                //        //changed on 02/25/2011
                //        //lochunk = new Chunk(lsFormatedString, Font9Bold());
                //        lochunk = new Chunk(lsFormatedString, Font8Bold());
                //        loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
                //        if (lsFormatedString.Length > 25)
                //        {
                //            if (checkTrue(lodataset, liRowCount, "_OrderNmb"))
                //            {
                //                //decrease columncount by 1 to adjust the Colspan. eg: NON-INVESTMENT ASSETS/LOOK-THROUGHS
                //                loCell.Colspan = 2;
                //                colsize = colsize - 1;
                //            }
                //        }
                //        setBottomWidthWhite(loCell);

                //    } /*=========IF END OF BOLD AND SUPERBOLD FLAG========*/
                //    else
                //    {
                //        if (lodataset.Tables[0].Rows[liRowCount][0].ToString() == "NET CHANGE")
                //        {
                //            setGreyBorder(loCell);
                //            //added on 28Feb2011 to change font size for total
                //            if (liColumnCount != 0)
                //            {
                //                lochunk = new Chunk(lsFormatedString, Font7Bold());
                //            }
                //        }
                //    }

                //    if (lodataset.Tables[0].Rows[liRowCount][0].ToString().Contains("NET CHANGE %"))
                //    {

                //        lsFormatedString = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]);
                //        try
                //        {
                //            lsFormatedString = String.Format("{0:#,###0.0%;(#,###0.0%)}", Convert.ToDecimal(lsFormatedString) / 100);
                //        }
                //        catch
                //        {

                //        }
                //        //changed on 02/25/2011
                //        //lochunk = new Chunk(lsFormatedString, Font9Bold());
                //        lochunk = new Chunk(lsFormatedString, Font8Bold());
                //        //added on 28Feb2011 to change font size for total
                //        if (liColumnCount != 0)
                //        {
                //            lochunk = new Chunk(lsFormatedString, Font7Bold());
                //        }


                //    }
                //    #endregion

                //}
                //else
                //{
                //    if (liColumnCount == 0 && !checkTrue(lodataset, liRowCount, "_OrderNmb"))
                //    {
                //        String abc = "" + lodataset.Tables[0].Rows[liRowCount][0].ToString();
                //        //changed on 02/25/2011
                //        //lochunk = new Chunk(abc, Font9Whitecheck(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][0])));
                //        lochunk = new Chunk(abc, Font7Whitecheck(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][0])));

                //        if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "")
                //        {
                //            //loCell.EnableBorderSide(0);
                //            lochunk = new Chunk(abc, Font18Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount]["Investment"])));
                //        }
                //        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "6")
                //        {
                //            checkTrue(lodataset, liRowCount, "_OrderNmb", loCell, new iTextSharp.text.Color(216, 216, 216));
                //            lochunk = new Chunk(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]), setFontsAll(8, 1, 0));
                //            //lochunk.SetBackground(iTextSharp.text.Color.LIGHT_GRAY);#B7DDE8 new iTextSharp.text.Color(216, 216, 216)
                //        }
                //        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "3")
                //        {

                //            if (liRowCount == lodataset.Tables[0].Rows.Count - 5)
                //            {
                //                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                //                loCell.VerticalAlignment = 4;
                //                loCell.Leading = 10f;
                //                lsFormatedString = "TOTAL PROPOSED/CONFIRMED COMMITMENTS ";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                lochunk = new Chunk(lsFormatedString, setFontsAll(8, 1, 0));
                //            }
                //            else if (liRowCount == lodataset.Tables[0].Rows.Count - 4)
                //            {
                //                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                //                loCell.VerticalAlignment = 4;
                //                loCell.Leading = 10f;
                //                lsFormatedString = "TOTAL PROPOSED/CONFIRMED COMMITMENTS ";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                lochunk = new Chunk(lsFormatedString, setFontsAll(8, 1, 0));
                //            }
                //            else //if (liRowCount == lodataset.Tables[0].Rows.Count - 2)
                //            {
                //                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                //                loCell.VerticalAlignment = 4;
                //                loCell.Leading = 10f;
                //                lsFormatedString = "TOTAL COMMITMENTS ";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                lochunk = new Chunk(lsFormatedString, setFontsAll(8, 1, 0));
                //            }

                //        }
                //        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "2")
                //        {
                //            if (liRowCount == lodataset.Tables[0].Rows.Count - 2)
                //            {
                //                lsFormatedString = "";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                lochunk = new Chunk(lsFormatedString, setFontsAll(7, 0, 0));
                //            }
                //            else
                //            {
                //                lsFormatedString = "";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                lochunk = new Chunk(lsFormatedString, setFontsAll(7, 0, 0));
                //            }

                //        }
                //        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "7")
                //        {
                //            if (liRowCount == lodataset.Tables[0].Rows.Count - 2)
                //            {
                //                loCell.VerticalAlignment = 4;
                //                loCell.Leading = 10f;
                //                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                //                lsFormatedString = "TOTAL COMMITMENTS ";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                lochunk = new Chunk(lsFormatedString, setFontsAll(8, 1, 0));
                //            }
                //            else
                //            {
                //                //loCell.EnableBorderSide(1);
                //                lsFormatedString = "";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                lochunk = new Chunk(lsFormatedString, setFontsAll(7, 1, 0));
                //            }

                //        }
                //        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "8")
                //        {
                //            loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                //            //lochunk = new Chunk(abc, Font18Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount]["Investment"])));
                //        }
                //    }
                //    else if (liColumnCount != 0 && !checkTrue(lodataset, liRowCount, "_OrderNmb"))
                //    {
                //        if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "2")
                //        {
                //            if (Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]) != "")
                //            {
                //                try
                //                {
                //                    lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                    lochunk = new Chunk(lsFormatedString, Font18Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                }
                //                catch
                //                {

                //                }
                //            }
                //        }
                //        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "3")
                //        {
                //            if (Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]) != "")
                //            {
                //                try
                //                {
                //                    loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                //                    lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                    lochunk = new Chunk(lsFormatedString, Font19Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                }
                //                catch
                //                {

                //                }
                //            }
                //        }
                //        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "6")
                //        {

                //            try
                //            {
                //                checkTrue(lodataset, liRowCount, "_OrderNmb", loCell, new iTextSharp.text.Color(216, 216, 216));
                //                lochunk = new Chunk(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]), setFontsAll(9, 1, 0));
                //                //lochunk.SetBackground(iTextSharp.text.Color.LIGHT_GRAY);
                //            }
                //            catch
                //            { }
                //        }
                //        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "7")
                //        {
                //            if (Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]) != "")
                //            {
                //                try
                //                {
                //                    //loCell.EnableBorderSide(1);

                //                    loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                //                    lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                    lochunk = new Chunk(lsFormatedString, Font18Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                }
                //                catch
                //                {

                //                }
                //            }
                //        }
                //        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "8")
                //        {
                //            loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                //            //lochunk = new Chunk(abc, Font18Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount]["Investment"])));
                //        }


                //    }



                //}
                //if (checkTrue(lodataset, liRowCount, "_OrderNmb") && !checkTrue(lodataset, liRowCount, "_OrderNmb"))
                //{
                //    if (liColumnCount == 0)
                //    {
                //        String abc = "          " + "          " + lodataset.Tables[0].Rows[liRowCount][0].ToString();
                //        //changed on 02/25/2011
                //        //lochunk = new Chunk(abc, Font8Grey());
                //        lochunk = new Chunk(abc, Font7Grey());
                //    }
                //    else
                //    {
                //        //changed on 02/25/2011
                //        //lochunk = new Chunk(lsFormatedString, Font8Grey());
                //        lochunk = new Chunk(lsFormatedString, Font7Grey());
                //    }
                //}

                ////CONDITION FOR SUPERBOLDFLAG
                ////checkTrue(lodataset, liRowCount, "_OrderNmb", loCell, new iTextSharp.text.Color(183, 221, 232));
                ////====added on 28Feb2011 to change font size for total====
                //if (checkTrue(lodataset, liRowCount, "_OrderNmb"))
                //{
                //    if (liColumnCount != 0)
                //    {
                //        lochunk = new Chunk(lsFormatedString, Font7Bold());
                //    }
                //}
                ///*=====END=====*/

                //if (checkTrue(lodataset, liRowCount, "_OrderNmb"))
                //{
                //    if (liColumnCount == 0)
                //    {
                //        String abc = "          " + "          " + "Total";
                //        //changed on 02/25/2011
                //        //lochunk = new Chunk(abc, Font8Normal());
                //        lochunk = new Chunk(abc, Font7Normal());
                //    }
                //    setTopWidthBlack(loCell);
                //    setBottomWidthWhite(loCell);

                //}
                #endregion
                loCell.Add(lochunk);
                loTable.AddCell(loCell);
            }

            try
            {
                if (liRowCount == loInsertdataset.Tables[0].Rows.Count - 1)
                {
                    document.Add(loTable);
                    liCurrentPage = liCurrentPage + 1;
                    document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, loInsertdataset.Tables[0].Rows.Count % liPageSize, true, lsFooterTxt));
                }
            }
            catch (Exception Ex)
            {

            }
        }

        if (loInsertdataset.Tables[0].Rows.Count > 0)
        {
            document.Close();

            FileInfo loFile = new FileInfo(ls);
            try
            {
                loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));

                Response.Write("<script>");
                string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + strGUID + ".pdf";
                Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
                Response.Write("</script>");

            }
            catch (Exception exc)
            {
                Response.Write(exc.Message);
            }
        }
    }
    
    public void setTopWidthBlack(Cell foCell)
    {
        foCell.BorderColor = iTextSharp.text.Color.BLACK;
        foCell.Border = iTextSharp.text.Rectangle.TOP_BORDER;
        foCell.BorderWidth = 0.1F;
    }
    public void setGreyBorder(DataSet foDataset, Cell foCell, int fiRowCount)
    {
        try
        {
            if (checkTrue(foDataset, fiRowCount, "_OrderNmb") || checkTrue(foDataset, fiRowCount, "_OrderNmb") || checkTrue(foDataset, fiRowCount, "_OrderNmb"))
            {
                setBottomWidthWhite(foCell);
            }
            if (checkTrue(foDataset, fiRowCount + 1, "_OrderNmb") || checkTrue(foDataset, fiRowCount + 1, "_OrderNmb") || checkTrue(foDataset, fiRowCount + 1, "_OrderNmb"))
            {
                setBottomWidthWhite(foCell);
            }
            else
            {
                foCell.BorderWidthBottom = 0.1F;
                foCell.BorderColorBottom = new iTextSharp.text.Color(216, 216, 216);
                //foCell.BorderColorBottom = new iTextSharp.text.Color(121, 121, 121);
            }
        }
        catch { }
    }

    public void setGreyBorder(Cell foCell)
    {

        foCell.BorderWidthBottom = 0.1F;
        //foCell.BorderColorBottom = new iTextSharp.text.Color(242, 242, 242);
        foCell.BorderColorBottom = new iTextSharp.text.Color(216, 216, 216);

    }
    public void setBottomWidthWhite(Cell foCell)
    {
        foCell.BorderWidthBottom = 0;
        foCell.BorderColorBottom = new iTextSharp.text.Color(255, 255, 255);
    }
    public iTextSharp.text.Font setFontsAll(int size, int bold, int italic, iTextSharp.text.Color foColor)
    {
        #region WITH OLD FONTS FROM FRUTIGER
        //string fontpath = Server.MapPath(".");
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        //if (bold == 1)
        //{
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD, foColor);
        //}
        //if (italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger_italic.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        //}
        //if (bold == 1 && italic == 1)
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC, foColor);
        //return font; 
        #endregion

        #region WITH NEW FONTS FROM FRUTIGER
        string fontpath = Server.MapPath(".");
        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTR_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        if (bold == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBL____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD, foColor);
        }
        if (italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTI_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        }
        if (bold == 1 && italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBLI___.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC, foColor);
        }
        return font;
        #endregion
    }
    public iTextSharp.text.Font setFontsAll(int size, int bold, int italic)
    {
        #region WITH OLD FONTS FROM FRUTIGER
        //string fontpath = Server.MapPath(".");
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //if (bold == 1)
        //{
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        //}
        //if (italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger_italic.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //}
        //if (bold == 1 && italic == 1)
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
        //return font; 
        #endregion

        #region WITH NEW FONTS FROM FRUTIGER
        string fontpath = Server.MapPath(".");
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTR_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        if (bold == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBL____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        }
        if (italic == 1)
        {
            //FTI_____.PFM
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTI_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        }
        if (bold == 1 && italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBLI___.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
        }

        return font;
        #endregion
    }


    public iTextSharp.text.Font setFontsAll1(int size, int bold)
    {
        #region WITH OLD FONTS FROM FRUTIGER
        //string fontpath = Server.MapPath(".");
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //if (bold == 1)
        //{
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        //}
        //if (italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger_italic.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //}
        //if (bold == 1 && italic == 1)
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
        //return font; 
        #endregion

        #region WITH NEW FONTS FROM FRUTIGER
        string fontpath = Server.MapPath(".");
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTR_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        if (bold == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBL____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.UNDERLINE);

            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        }


        return font;
        #endregion
    }


    public void setHeader(Document foDocument, DataSet loInsertdataset, DataSet loDatatset)
    {
       
        DataSet OldDataset = new DataSet();
        OldDataset = loDatatset.Copy();

        for (int liNewdataset = loInsertdataset.Tables[0].Columns.Count - 1; liNewdataset > -1; liNewdataset--)
        {
            if (loInsertdataset.Tables[0].Columns[liNewdataset].ColumnName.Contains("_") || loInsertdataset.Tables[0].Columns[liNewdataset].ColumnName.Trim().Equals("1"))
            {
                loInsertdataset.Tables[0].Columns.Remove(loInsertdataset.Tables[0].Columns[liNewdataset]);
            }
        }

        DataSet AddDataset = loInsertdataset.Copy();
        AddDataset.AcceptChanges();
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(6);   // 2 rows, 2 columns        
        setTableProperty(loTable);
        Chunk loParagraph = new Chunk();


        //////// set header new addition for pdf
        string lsfamilyName = "";
        if (ddlHousehold.SelectedValue != "0")
        {
            if (drpAllocationGroupTitle.SelectedValue == "0" && ddlAllocationGroup.SelectedValue != "0")
            {
                lsfamilyName = ddlAllocationGroup.SelectedItem.Text;
            }
            else if (ddlHousehold.SelectedValue != "0" && ddlAllocationGroup.SelectedValue == "0")
            {
                lsfamilyName = drpHouseHoldReportTitle.SelectedItem.Text;
            }
            else
            {
                lsfamilyName = drpAllocationGroupTitle.SelectedItem.Text;
            }
        }




        if (txtAsofdate.Text != "")
            lsDateName = Convert.ToDateTime(txtAsofdate.Text).ToString("MMMM dd, yyyy") + "";

        /////////////

        //Chunk lochunk = new Chunk(lsFamiliesName, iTextSharp.text.FontFactory.GetFont("frutigerce-roman", BaseFont.CP1252, BaseFont.EMBEDDED, 14, iTextSharp.text.Font.BOLD));

        Chunk lochunk = new Chunk(lsfamilyName, setFontsAll(12, 1, 0));
        iTextSharp.text.Cell loCell = new Cell();
        loCell.Add(lochunk);

        lochunk = new Chunk("\n" + "ASSET ALLOCATION SUMMARY", setFontsAll(11, 0, 0));
        //loParagraph.Chunks.Add(lochunk);

        loCell.Add(lochunk);
        loCell.Colspan = loInsertdataset.Tables[0].Columns.Count;
        loCell.HorizontalAlignment = 1;



        lochunk = new Chunk("\n" + lsDateName, setFontsAll(8, 0, 1)); //To Show date in header uncomment this
        loCell.Add(lochunk);
        loCell.Border = 0;
        //   loCell.Add(loParagraph);
        loTable.AddCell(loCell);

        Boolean lbCheckFoMarket = false;
        #region No Use
        //for (int liColumnCount = 0; liColumnCount < AddDataset.Tables[0].Columns.Count; liColumnCount++)
        //{

        //    if (Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName) != "")
        //    {
        //        if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("Current Portfolio Value"))
        //        {
        //            lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("Current Portfolio Value", "Current \n $"), setFontsAll1(7, 1));
        //        }
        //        else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("Current Portfolio %"))
        //        {
        //            lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("Current Portfolio %", ""), setFontsAll1(7, 1));
        //        }
        //        else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("Suggested Allocation"))
        //        {
        //            lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("Suggested Allocation", "Strategic Allocation"), setFontsAll(7, 1, 0));
        //        }
        //        else
        //        {
        //            lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName), setFontsAll(7, 1, 0));
        //        }
        //    }
        //    //}
        //    loCell = new Cell();

        //    loCell.Add(lochunk);
        //    loCell.Border = 0;
        //    loCell.NoWrap = true;//true;

        //    loCell.MaxLines = 2;
        //    loCell.Leading = -2F;
        //    if (liColumnCount == 0 )
        //    {
        //        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        //        loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        //    }

        //    else if (liColumnCount == 1)
        //    {
        //        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
        //        loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        //        loCell.BorderWidthBottom = 1f;
        //    }
        //    else
        //    {
        //        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        //        loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        //    }


        //    if (Convert.ToString(loInsertdataset.Tables[0].Columns[liColumnCount].ColumnName).Contains(" "))
        //    {
        //        loCell.Leading = 10f;//8
        //        loCell.MaxLines = 5;
        //        //loCell.Leading = 9f;
        //    }
        //    loCell.Leading = 10f;//8

        //    loCell.VerticalAlignment = 1; //5 ,6 bottom : WASTE VALUES - 3,4
        //    loTable.AddCell(loCell);

        //}
#endregion

        lochunk = new Chunk("", setFontsAll(7, 1,0));
        iTextSharp.text.Cell loCell0 = new Cell();
        loCell0.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        loCell0.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell0.Border = 0;
        loTable.AddCell(loCell0);

        lochunk = new Chunk("Current  Allocation", setFontsAll1(7, 0));
        iTextSharp.text.Chunk lochunk5 = new Chunk("\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "$                         %", setFontsAll(7, 1, 0));
        iTextSharp.text.Chunk lochunk1 = new Chunk(lochunk + "\n ", setFontsAll(7, 1, 0));
        
          
        iTextSharp.text.Cell loCell1 = new Cell();
        
        loCell1.Add(lochunk1);//.SetUnderline(0.8f,-1f)
        
        loCell1.Add(lochunk5);

        loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        loCell1.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell1.Colspan = 2;
        loCell1.Border = 0;
        loCell1.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
        //loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
        loCell1.MaxLines = 3;
        loCell1.Leading = 10f;
        //loCell1.EnableBorderSide(2);
        loTable.AddCell(loCell1);

        iTextSharp.text.Chunk lochunk2 = new Chunk("Strategic " + "\n ", setFontsAll(7, 1, 0));
        iTextSharp.text.Chunk lochunk6 = new Chunk("Allocation", setFontsAll(7, 1, 0));
        iTextSharp.text.Cell loCell2 = new Cell();
        loCell2.Add(lochunk2);
        loCell2.Add(lochunk6);
        loCell2.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        loCell2.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell2.Border = 0;
        loCell2.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
        loCell2.MaxLines = 3;
        loCell2.Leading = 10f;
        loTable.AddCell(loCell2);

        iTextSharp.text.Chunk lochunk3 = new Chunk("Tactical " + "\n ", setFontsAll(7, 1, 0));
        iTextSharp.text.Chunk lochunk7 = new Chunk("Tilt", setFontsAll(7, 1, 0));
        iTextSharp.text.Cell loCell3 = new Cell();
        loCell3.Add(lochunk3);
        loCell3.Add(lochunk7);
        loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        loCell3.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell3.Border = 0;
        //loCell3.EnableBorderSide(2);
        loCell3.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
        //loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
        loCell3.MaxLines = 3;
        loCell3.Leading = 10f;
        loTable.AddCell(loCell3);

        iTextSharp.text.Chunk lochunk4 = new Chunk("Tactical " + "\n ", setFontsAll(7, 1, 0));
        iTextSharp.text.Chunk lochunk8 = new Chunk("Target", setFontsAll(7, 1, 0));

        iTextSharp.text.Cell loCell4 = new Cell();
        loCell4.Add(lochunk4);
        loCell4.Add(lochunk8);

        loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        loCell4.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell4.Border = 0;
        //loCell4.EnableBorderSide(2);
        loCell4.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
        //loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
        loCell4.MaxLines = 3;
        loCell4.Leading = 10f;
        
        loTable.AddCell(loCell4);


        foDocument.Add(loTable);

        //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(@"C:\AdventReport\images\Gresham_Logo.png");
        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        foDocument.Add(png);
    }

    public void setTableProperty(iTextSharp.text.Table fotable)
    {
        //int[] headerwidths = { 28, 9, 9, 9, 9, 9, 9, 9, 7 };

        setWidthsoftable(fotable);

        //fotable.Width = 100;
        fotable.Alignment = 1;
        fotable.Border = 0;
        fotable.Cellspacing = 0;
        fotable.Cellpadding = 3;
        fotable.Locked = false;

    }

    public void setWidthsoftable(iTextSharp.text.Table fotable)
    {

        switch (lsTotalNumberofColumns)
        {
            case "2":
                int[] headerwidths2 = { 30, 9 };
                fotable.SetWidths(headerwidths2);
                fotable.Width = 40;
                break;
            case "3":
                int[] headerwidths3 = { 30, 9, 9 };
                fotable.SetWidths(headerwidths3);
                fotable.Width = 49;
                break;
            case "4":
                int[] headerwidths4 = { 15, 13, 13, 16 };
                fotable.SetWidths(headerwidths4);
                fotable.Width = 53;
                break;
            case "5":
                int[] headerwidths5 = { 30, 9, 9, 9, 9 };
                fotable.SetWidths(headerwidths5);
                fotable.Width = 67;
                break;
            case "6":
                int[] headerwidths6 = { 27, 11, 11, 8, 5, 7 };
                fotable.SetWidths(headerwidths6);
                fotable.Width = 70;
                break;
            case "7":
                int[] headerwidths7 = { 30, 9, 9, 9, 9, 9, 9 };
                fotable.SetWidths(headerwidths7);
                fotable.Width = 85;
                break;
            case "8":
                int[] headerwidths8 = { 30, 9, 9, 9, 9, 9, 9, 9 };
                fotable.SetWidths(headerwidths8);
                fotable.Width = 94;
                break;
            case "9":
                int[] headerwidths9 = { 27, 9, 9, 9, 9, 9, 9, 9, 7 };
                fotable.SetWidths(headerwidths9);
                fotable.Width = 97;
                break;

            case "10":
                int[] headerwidths10 = { 25, 8, 8, 8, 8, 8, 8, 8, 8, 8 };
                fotable.SetWidths(headerwidths10);
                fotable.Width = 97; break;
            case "11":
                //int[] headerwidths11 = { 25, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7 };
                int[] headerwidths11 = { 20, 8, 8, 8, 12, 20, 8, 8, 8, 8, 8 };
                fotable.SetWidths(headerwidths11);
                fotable.Width = 98; break;
            case "12":
                int[] headerwidths12 = { 25, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7 };
                fotable.SetWidths(headerwidths12);
                fotable.Width = 102; break;
            case "13":
                int[] headerwidths13 = { 30, 9 };
                fotable.SetWidths(headerwidths13);
                fotable.Width = 39; break;
            case "14":
                int[] headerwidths14 = { 30, 9 };
                fotable.SetWidths(headerwidths14);
                fotable.Width = 39;
                break;
            case "15":
                int[] headerwidths15 = { 30, 9 };
                fotable.SetWidths(headerwidths15);
                fotable.Width = 39;
                break;
            case "16":
                int[] headerwidths16 = { 30, 9 };
                fotable.SetWidths(headerwidths16);
                fotable.Width = 39;
                break;
            case "17":
                int[] headerwidths17 = { 30, 9 };
                fotable.SetWidths(headerwidths17);
                fotable.Width = 39;
                break;
            case "18":
                int[] headerwidths18 = { 30, 9 };
                fotable.SetWidths(headerwidths18);
                fotable.Width = 39;
                break;
            case "19":
                int[] headerwidths19 = { 30, 9 };
                fotable.SetWidths(headerwidths19);
                fotable.Width = 39;
                break;
            case "20":
                int[] headerwidths20 = { 30, 9 };
                fotable.SetWidths(headerwidths20);
                fotable.Width = 39;
                break;

        }
    }

    public Boolean checkTrue(DataSet foDataset, int fiRowCount, String fsField)
    {
        Boolean lblReturn = false;
        if (foDataset.Tables[0].Rows.Count > 0)
        {
            if (foDataset.Tables[0].Rows[fiRowCount][fsField].ToString().ToUpper() == "TRUE")
            {
                lblReturn = true;
            }
        }
        return lblReturn;

    }
    public iTextSharp.text.Font Font9Normal()
    {
        return setFontsAll(9, 0, 0);
    }
    public iTextSharp.text.Font Font1Normal()
    {
        return setFontsAll(1, 0, 0);
    }
    public iTextSharp.text.Font Font8Normal()
    {
        return setFontsAll(8, 0, 0);
    }

    public iTextSharp.text.Font Font7Normal()
    {
        return setFontsAll(7, 0, 0);
    }

    public iTextSharp.text.Font Font8GreyItalic()
    {
        return setFontsAll(8, 0, 1, new iTextSharp.text.Color(216, 216, 216));
    }

    public iTextSharp.text.Font Font7GreyItalic()
    {
        return setFontsAll(7, 0, 1, new iTextSharp.text.Color(216, 216, 216));
    }
    public iTextSharp.text.Font Font8Grey()
    {
        return setFontsAll(8, 0, 0, new iTextSharp.text.Color(175, 175, 175));
        //return setFontsAll(9, 0, 0, new iTextSharp.text.Color(175, 175, 175));
    }

    public iTextSharp.text.Font Font7Grey()
    {
        //return setFontsAll(7, 0, 0, new iTextSharp.text.Color(175, 175, 175));
        return setFontsAll(7, 0, 0, new iTextSharp.text.Color(165, 165, 165));
        //return setFontsAll(9, 0, 0, new iTextSharp.text.Color(175, 175, 175));
    }

    public iTextSharp.text.Font Font8Whitecheck(String fsTest)
    {
        if (fsTest == "test")
            return setFontsAll(8, 0, 0, new iTextSharp.text.Color(255, 255, 255));
        else
            return setFontsAll(8, 0, 0);
    }

    public iTextSharp.text.Font Font7Whitecheck(String fsTest)
    {
        if (fsTest == "test")
            return setFontsAll(7, 0, 0, new iTextSharp.text.Color(255, 255, 255));
        else
            return setFontsAll(7, 0, 0);
    }

    public iTextSharp.text.Font Font9Whitecheck(String fsTest)
    {
        if (fsTest == "test")
            return setFontsAll(9, 0, 0, new iTextSharp.text.Color(255, 255, 255));
        else
            return setFontsAll(9, 0, 0);
    }
    public iTextSharp.text.Font Font9Bold()
    {
        return setFontsAll(9, 1, 0);
    }

    public iTextSharp.text.Font Font8Bold()
    {
        return setFontsAll(8, 1, 0);
    }

    public iTextSharp.text.Font Font18Bold(String fsTest)
    {
        if (fsTest == "test")
        {
            return setFontsAll(7, 1, 0);
        }
        else
        {
            return setFontsAll(7, 1, 0);
        }
    }

    public iTextSharp.text.Font Font19Bold(String fsTest)
    {
        if (fsTest == "test")
        {
            return setFontsAll1(7, 1);
        }
        else
        {
            return setFontsAll1(7, 1);
        }
    }

    public iTextSharp.text.Font Font7Bold()
    {
        return setFontsAll(7, 1, 0);
    }

    public void checkTrue(DataSet foDataset, int fiRowCount, String fsField, Cell foCell, iTextSharp.text.Color foColor)
    {

        if (foDataset.Tables[0].Rows[fiRowCount][fsField].ToString() == "6")
        {
            foCell.BackgroundColor = foColor;
            foCell.VerticalAlignment = 4;
            foCell.Leading = 10f;
        }


    }
    public iTextSharp.text.Table addFooter(String lsDateTime, int liTotalPages, int liCurrentPage, int liLastPageData, Boolean footerflg, String FooterTxt)
    {

        iTextSharp.text.Table fotable = new iTextSharp.text.Table(2, 1);
        fotable.Width = 90;
        fotable.Border = 0;
        int[] headerwidths = { 50, 40 };
        fotable.SetWidths(headerwidths);
        fotable.Cellpadding = 0;
        Cell loCell = new Cell();
        Chunk loChunk = new Chunk();

        for (int liCounter = 0; liCounter < liPageSize - 2 - liLastPageData; liCounter++)
        {
            loCell = new Cell();
            loChunk = new Chunk("dev", Font8Whitecheck("test"));
            loCell.HorizontalAlignment = 2;
            loCell.BorderWidth = 0;
            loCell.Add(loChunk);
            fotable.AddCell(loCell);

            loCell = new Cell();
            loChunk = new Chunk("dev", Font8Whitecheck("test"));
            loCell.Add(loChunk);
            loCell.BorderWidth = 0;
            loCell.HorizontalAlignment = 2;
            fotable.AddCell(loCell);
        }
        if (footerflg)
        {
            loCell = new Cell();
            //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
            loChunk = new Chunk(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
            //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
            loCell.Leading = 8f;
            loCell.HorizontalAlignment = 0;
            loCell.Colspan = 2;
            loCell.BorderWidth = 0;
            loCell.Add(loChunk);
            fotable.AddCell(loCell);
        }


        /* Please uncomment this to show page numbers */

        //loCell = new Cell();
        //loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font8Normal());
        //loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font7Normal());
        //loCell.Leading = 15f;//25f
        //loCell.HorizontalAlignment = 2;
        //loCell.BorderWidth = 0;
        //loCell.Add(loChunk);
        //fotable.AddCell(loCell);

        loCell = new Cell();
        //loChunk = new Chunk(lsDateTime, Font8GreyItalic());
        loChunk = new Chunk(lsDateTime, Font7GreyItalic());
        loCell.Add(loChunk);
        loCell.Leading = 15f;//25f
        loCell.BorderWidth = 0;
        loCell.HorizontalAlignment = 2;
        fotable.AddCell(loCell);
        //fotable.TableFitsPage = true;

        return fotable;
    }

    protected void ddlAllocationGroup_SelectedIndexChanged(object sender, EventArgs e)
    {
        fillGroupAllocationTitle();
    }
}
