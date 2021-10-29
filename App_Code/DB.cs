using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
//using CrmSdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
/// <summary>
/// Summary description for DB
/// </summary>
public class DB
{
    //GeneralMethods clsGM = new GeneralMethods();
	public DB()
	{
		// constructor dont delete
    }

    SqlConnection  cn;
	SqlCommand cmd;
	SqlDataReader dr;
    DataSet ds_Configuration=new DataSet();

  //  String lsConnectionstring = System.Configuration.ConfigurationManager.ConnectionStrings["connectotsql"].ConnectionString;

String lsConnectionstring = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);//"Password=slater6;Persist Security //Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";

  //  String lsConnectionstring = System.Configuration.ConfigurationManager.ConnectionStrings["connectotsqllocal"].ConnectionString;
    
    public String _dbErrorMsg;
    public String _ProceduerOutPara;
    public String _ProceduerReturnPara;

    public string ConnString = "";
    public string TempOutDir = "";
    public string DocumentOutDir = "";
        
    public void close()
    { 
       
    }
    // To Open Database Connection
    public void getConfiguration()
    {
        ds_Configuration.ReadXml(AppDomain.CurrentDomain.BaseDirectory + "\\Configuration.xml");
        if (ds_Configuration.Tables.Count != 0)
        {
            if (ds_Configuration.Tables["filepaths"].Rows.Count != 0)
            {
                foreach (DataRow drXMLRule in ds_Configuration.Tables["filepaths"].Rows)
                {
                    TempOutDir = drXMLRule[0].ToString();
                    DocumentOutDir = drXMLRule[1].ToString();
                }
            }
        }
    }

  public DataSet ExecuteSPOutParameter(string cmdText, int TimeOut, params SqlParameter[] cmdParams)
    {
        //Create a connection to the SQL Server; modify the connection string for your environment.
        SqlConnection conn = new SqlConnection(lsConnectionstring);

        //Create a DataAdapter, and then provide the name of the stored procedure.
        SqlDataAdapter objda = new SqlDataAdapter(cmdText, conn);
        objda.SelectCommand.CommandTimeout = TimeOut;
        //Set the command type as StoredProcedure.
        objda.SelectCommand.CommandType = CommandType.StoredProcedure;

        if (cmdParams != null)
        {
            foreach (SqlParameter objParm in cmdParams)
                objda.SelectCommand.Parameters.Add(objParm);
        }
        //Create a new DataSet to hold the records.
        DataSet objDS = new DataSet();

        objda.Fill(objDS);

        return objDS;
    }
    public string ParseDate(string Date)
    {
        string DD, MM, yy;

        if (Date.Contains("-") || Date.Contains("/") || Date.Length == 0)
            return Date.Replace("-", "/");
        else
        {
            switch (Date.Length)
            {
                case 4:
                    MM = Date.Substring(0, 1);
                    DD = Date.Substring(1, 1);
                    yy = Date.Substring(2);
                    Date = MM + "/" + DD + "/" + yy;
                    break;
                case 6:
                    if (Date.StartsWith("0") || Date.StartsWith("10") || Date.StartsWith("11") || Date.StartsWith("12"))
                    {
                        MM = Date.Substring(0, 2);
                        DD = Date.Substring(2, 2);
                        yy = Date.Substring(4);
                        Date = MM + "/" + DD + "/" + yy;
                    }
                    else
                    {
                        MM = Date.Substring(0, 1);
                        DD = Date.Substring(1, 1);
                        yy = Date.Substring(2);
                        Date = MM + "/" + DD + "/" + yy;
                    }
                    break;
                case 8:
                    MM = Date.Substring(0, 2);
                    DD = Date.Substring(2, 2);
                    yy = Date.Substring(4);
                    Date = MM + "/" + DD + "/" + yy;
                    break;
                default:
                    Date = Date;
                    break;
            }
            return Date;

        }
    }
    public SqlConnection gOpenConnection()
    {   try
        {
            ConnString = lsConnectionstring;
            cn = new SqlConnection(ConnString);
            cn.Open();
            return cn;
        }
        catch(Exception ex)
        {
           _dbErrorMsg = ex.Message;
           return cn;
        }
    }
    
    // To close Database Connection
    public void gCloseConnection()
    {
        cn.Close();
        cn.Dispose();
    }
    public void gCloseConnection(SqlConnection cn)
    {
        cn.Close();
        cn.Dispose();
    }
    // To add record
	public String AddRecord(string sqlstr,string ExecuteType)
	{
        try
        {
            string primaryKeyValue;
            cn = gOpenConnection();

            if (ExecuteType == "ID")
            {
                sqlstr = "SET NOCOUNT ON " + sqlstr.ToString().Trim() + " SELECT @@IDENTITY SET NOCOUNT OFF ";
            }
            cmd = new SqlCommand(sqlstr, cn);
            if (cn.State.ToString() == "Open")
            {
                try
                {
                    cmd.CommandText = sqlstr;
                    if (ExecuteType.ToString() == "ID")
                    {
                        primaryKeyValue = cmd.ExecuteScalar().ToString();
                        //gCloseConnection(cn);
                        return primaryKeyValue;
                    }
                    else
                    {
                        cmd.ExecuteNonQuery();
                        return "";//New record has been Added.
                    }
                }
                catch (Exception e)
                {
                    return e.ToString();
                }
            }
            else
            {
                return cn.State.ToString();
            }
        }
        finally
        {
            if (cn != null)
                gCloseConnection(cn);
        }
	}
    // To Update Record
	public String UpdateRecord(String sqlstr)
	{
        cn=gOpenConnection();
        cmd = new SqlCommand(sqlstr, cn);
		if (cn.State.ToString()=="Open")
		{
			try
			{
				cmd.ExecuteNonQuery();
                gCloseConnection(cn);
                return "";//Record has been Updated
			}
			catch(Exception e)
			{
				return e.ToString();
			}
		}
		else
		{
			return cn.State.ToString();
		}
		
	}
    // To Delete Record
	public String DeleteRecord(String sqlstr)
	{
        cn=gOpenConnection();
        cmd = new SqlCommand(sqlstr, cn);
		if (cn.State.ToString()=="Open")
		{
			try
			{
				cmd.ExecuteNonQuery();
                gCloseConnection(cn);
                return "";//Record has been Deleted
			}
			catch(Exception e)
			{
				return e.ToString();
			}
		}
		else
		{
			return cn.State.ToString();
		}
		
	}

      // To Delete Record
   // public String DeleteRecord(String sqlstr, EntityName entityName, CrmService crmService)
      public String DeleteRecord(String sqlstr, string entityName,IOrganizationService crmService)
    {
        //IOrganizationService service = null;
        int num2 = 0;
        try
        {
            DataSet dataSet = getDataSet(sqlstr);
            for (int j = 0; j < dataSet.Tables[0].Rows.Count; j++)
            {
                Guid UUID = new Guid(Convert.ToString(dataSet.Tables[0].Rows[j][0]));
                //crmService.Delete(entityName.ToString(), UUID);
               //service= clsGM.GetCrmService();
               crmService.Delete(entityName, UUID);
                num2 = num2 + 1;
            }

            return num2.ToString();//Record has been Deleted
        }
        catch (Exception e)
        {
            return e.ToString();
        }

    }

    // Fill Dataset
	public DataSet getDataSet(String sqlstr)
	{
        cn=gOpenConnection();        
		SqlDataAdapter da = new SqlDataAdapter(sqlstr,cn);        
        da.SelectCommand.CommandTimeout = 2400;
		DataSet ds = new DataSet();
		da.Fill(ds);
		da.Dispose();
        gCloseConnection(cn);
		return (ds);		
	}

    // Fill Dataset
    public DataSet getDataSet(String sqlstr,int TimeOut)
    {
        cn = gOpenConnection();
        SqlCommand objcmd = new SqlCommand(sqlstr, cn);
        objcmd.CommandTimeout = TimeOut;
        SqlDataAdapter da = new SqlDataAdapter(objcmd);
        DataSet ds = new DataSet();
        da.Fill(ds);
        da.Dispose();
        gCloseConnection(cn);
        return (ds);
    }
    // Fill Datareader
    //public SqlDataReader  getDataReader(String sqlstr)
    //{	
    //    if (sqlstr.ToString().Trim() != "")
    //    {
    //        cn=gOpenConnection();
    //        cmd = new SqlCommand(sqlstr,cn);
    //        dr=cmd.ExecuteReader();
    //        return (dr);
    //    }
    //    else
    //    {
    //        return (dr);
    //    }
    //}
    // Function to Check whether field/ condition in table exists
    //public bool checkRecord(String sqlstr)
    //{	
    //    if (sqlstr.ToString().Trim() != "")
    //    {
    //        cn=gOpenConnection();
    //        cmd = new SqlCommand(sqlstr,cn);
    //        dr=cmd.ExecuteReader();
    //        if  (dr.Read ())
    //        {
    //            dr.Close();
    //            gCloseConnection(cn);
    //            return (true);
    //        }
    //        else
    //        {
    //            return(false);
    //        }
    //    }
    //    else
    //    {
    //        return (false);
    //    }
    //}
    // Execute Sql Server Stored Procedure which returns/ do not return value
    public string Exec_Stored_procedure(String sqlstr,String exectype)
    {
        try
        {
            if (sqlstr.ToString().Trim() != "")
            {
                cn = gOpenConnection();
                cmd = new SqlCommand(sqlstr, cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = sqlstr;
                if (exectype == "ID")
                {
                    int procstatus = 0;
                    sqlstr = "SET NOCOUNT ON declare @ReturnStatus int exec @ReturnStatus = " + sqlstr.ToString().Trim() + " SELECT @ReturnStatus SET NOCOUNT OFF ";
                    cmd = new SqlCommand(sqlstr, cn);
                    procstatus = Convert.ToInt32(cmd.ExecuteScalar().ToString());
                    gCloseConnection(cn);
                    return procstatus.ToString();
                }
                else if (exectype == "SQ")
                {
                    int procstatus = 0;
                    sqlstr = "SET NOCOUNT ON " + sqlstr.ToString().Trim() + " SET NOCOUNT OFF ";
                    cmd = new SqlCommand(sqlstr, cn);
                    procstatus = Convert.ToInt32(cmd.ExecuteScalar().ToString());
                    gCloseConnection(cn);
                    return procstatus.ToString();
                }
                else
                {
                    //cn = gOpenConnection();
                    sqlstr = "exec " + sqlstr.ToString().Trim();
                    cmd = new SqlCommand(sqlstr, cn);
                    cmd.ExecuteNonQuery();
                    gCloseConnection(cn);
                    return "";
                }
            }
            else
            {
                return "false";
            }
        }
        finally
        {
            if (cn != null)
                gCloseConnection(cn);
        }
    }
    //check
    public string getValue(String columnName,String tableName,String WhereClause)
    {
        String sqlstr;
        String tempcolumn;   

        if (columnName=="" || tableName=="")
        {return "INVALID PARAMETERS";}
        else
        {
            if (WhereClause != "")
            { sqlstr = "Select " + columnName + " from " + tableName + " where " + WhereClause; }
            else
            { sqlstr = "Select " + columnName + " from " + tableName ; }

            if (sqlstr.ToString().Trim() != "")
            {
                cn = gOpenConnection();
                cmd = new SqlCommand(sqlstr, cn);
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    tempcolumn = dr[0].ToString();
                    dr.Close();
                    gCloseConnection(cn);
                    return tempcolumn;
                }
                else
                {
                    return "";
                }
            }
            else
            {
                return "";
            }
        }
    }
    public DataSet ExecuteSPOutParameter(string cmdText, int TimeOut, string outParamName, out string outParam, params SqlParameter[] cmdParams)
    {
        //Create a connection to the SQL Server; modify the connection string for your environment.
        SqlConnection conn = new SqlConnection(lsConnectionstring);

        //Create a DataAdapter, and then provide the name of the stored procedure.
        SqlDataAdapter objda = new SqlDataAdapter(cmdText, conn);
        objda.SelectCommand.CommandTimeout = TimeOut;
        //Set the command type as StoredProcedure.
        objda.SelectCommand.CommandType = CommandType.StoredProcedure;

        if (cmdParams != null)
        {
            foreach (SqlParameter objParm in cmdParams)
                objda.SelectCommand.Parameters.Add(objParm);
        }        
        //Create a new DataSet to hold the records.
        DataSet objDS = new DataSet();

        objda.Fill(objDS);
        outParam = Convert.ToString(objda.SelectCommand.Parameters[outParamName].Value);

        return objDS;
    }

    public object ExecuteScalar(string cmdText, string commandType, string outParamName, out string outParam, params SqlParameter[] cmdParams)
    {
        SqlConnection conn = null;
        try
        {
            conn = new SqlConnection(lsConnectionstring);
            conn.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            cmd.CommandText = cmdText;
            if (commandType == "StoredProcedure")
                cmd.CommandType = CommandType.StoredProcedure;
            else
                cmd.CommandType = CommandType.Text;
            // Add the command parameters to the command object.
            if (cmdParams != null)
            {
                foreach (SqlParameter objParm in cmdParams)
                    cmd.Parameters.Add(objParm);
            }
            // Execute command 
            Object _returnObj = cmd.ExecuteScalar();
            outParam = cmd.Parameters[outParamName].Value.ToString();
            cmd.Parameters.Clear();
            cmd.Dispose();
            return _returnObj;
        }
        finally
        {
            if (conn != null)
                conn.Close();
        }
    }
    public object ExecuteScalar(string cmdText, string commandType, string outParamName, out string outParam, string outParamName1, out string outParam1, params SqlParameter[] cmdParams)
    {
        SqlConnection conn = null;
        try
        {
            conn = new SqlConnection(lsConnectionstring);
            conn.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            cmd.CommandText = cmdText;
            if (commandType == "StoredProcedure")
                cmd.CommandType = CommandType.StoredProcedure;
            else
                cmd.CommandType = CommandType.Text;
            // Add the command parameters to the command object.
            if (cmdParams != null)
            {
                foreach (SqlParameter objParm in cmdParams)
                    
                    cmd.Parameters.Add(objParm);
            }
            // Execute command 
            Object _returnObj = cmd.ExecuteScalar();
            outParam = cmd.Parameters[outParamName].Value.ToString();
            outParam1 = cmd.Parameters[outParamName1].Value.ToString();
            cmd.Parameters.Clear();
            cmd.Dispose();
            return _returnObj;
        }
        finally
        {
            if (conn != null)
                conn.Close();
        }
    }

    public int ExecuteScalar(string cmdText, string commandType, params SqlParameter[] cmdParams)
    {
        SqlConnection conn = null;
        try
        {
            conn = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["TSI_UMS_SQLServer"].ConnectionString);
            conn.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            cmd.CommandText = cmdText;
            if (commandType == "StoredProcedure")
                cmd.CommandType = CommandType.StoredProcedure;
            else
                cmd.CommandType = CommandType.Text;
            // Add the command parameters to the command object.
            if (cmdParams != null)
            {
                foreach (SqlParameter objParm in cmdParams)
                    cmd.Parameters.Add(objParm);
            }
            // Execute command 
            int _returnObj = cmd.ExecuteNonQuery();
            //outParam = 0;
            cmd.Parameters.Clear();
            cmd.Dispose();
            return _returnObj;
        }
        finally
        {
            if (conn != null)
                conn.Close();
        }
    }

    public class DataSetHelper
    {
        public DataSet ds;
        public DataSetHelper(ref DataSet DataSet)
        {
            ds = DataSet;
        }
        public DataSetHelper()
        {
            ds = null;
        }
    }
    private bool ColumnEqual(object A, object B)
    {

        // Compares two values to see if they are equal. Also compares DBNULL.Value.
        // Note: If your DataTable contains object fields, then you must extend this
        // function to handle them in a meaningful way if you intend to group on them.

        if (A == DBNull.Value && B == DBNull.Value) //  both are DBNull.Value
            return true;
        if (A == DBNull.Value || B == DBNull.Value) //  only one is DBNull.Value
            return false;
        return (A.Equals(B));  // value type standard comparison
    }
    public DataTable SelectDistinct(string TableName, DataTable SourceTable, string FieldName1, string FieldName2, string FieldName3)
    {
        DataTable dt = new DataTable(TableName);
        dt.Columns.Add(FieldName1, SourceTable.Columns[FieldName1].DataType);
        dt.Columns.Add(FieldName2, SourceTable.Columns[FieldName2].DataType);
        dt.Columns.Add(FieldName3, SourceTable.Columns[FieldName3].DataType);
        DataTable dt2 = new DataTable(TableName);
        dt2.Columns.Add(FieldName1, SourceTable.Columns[FieldName1].DataType);
        dt2.Columns.Add(FieldName2, SourceTable.Columns[FieldName2].DataType);
        dt2.Columns.Add(FieldName3, SourceTable.Columns[FieldName3].DataType);
        object LastValue = null;
        foreach (DataRow dr in SourceTable.Select("", FieldName1))
        {
            if (LastValue == null || !(ColumnEqual(LastValue, dr[FieldName1])))
            {
                LastValue = dr[FieldName1];
                dt.Rows.Add(new object[] { LastValue, dr[FieldName2], dr[FieldName3] });
            
            }
        }
        foreach (DataRow dr in dt.Select("", FieldName2))
        {
            dt2.Rows.Add(new object[] { dr[FieldName1], dr[FieldName2], dr[FieldName3] });

        }
        
        return dt2;
    }
   
   
    #region returns State using SP_S_state_lkup proc.
    public void getStateList(DropDownList drpState, int CountryId)
    {
        string sqlstr = string.Empty;
        DataSet DS;
        //gOpenConnection();
        sqlstr = " SP_S_state_lkup " + CountryId;
        DS = getDataSet(sqlstr);
        if (DS.Tables[0].Rows.Count > 0)
        {
            DataTable dt = new DataTable();
            dt = DS.Tables[0];
            drpState.DataSource = dt;
            drpState.DataTextField = "NameTxt";
            drpState.DataValueField = "IdNmb";
            drpState.DataBind();
        }
        else
        drpState.SelectedIndex = drpState.Items.Count - 1;
        ListItem li = new ListItem();
        li.Text = "Select";
        li.Value = "0";
        drpState.Items.Add(li);
        drpState.SelectedValue = "0";
        //gCloseConnection();
    }
    #endregion

    #region returns Country using SP_S_cntry_lkup proc.
    public void getCountryList(DropDownList drpCntry)
    {
        string sqlstr = string.Empty;
        DataSet DS;
        cn = gOpenConnection();
        sqlstr = "SP_S_cntry_lkup";
        DS = getDataSet(sqlstr);
        if (DS.Tables[0].Rows.Count > 0)
        {
            DataTable dt = new DataTable();
            dt = DS.Tables[0];
            drpCntry.DataSource = dt;
            drpCntry.DataTextField = "NameTxt";
            drpCntry.DataValueField = "IdNmb";
            drpCntry.DataBind();
            drpCntry.Items.Insert(0, "Select");
            drpCntry.Items[0].Value = "0";
            drpCntry.SelectedIndex = 0;          

        }
        else
            drpCntry.SelectedIndex = drpCntry.Items.Count - 1;
        ListItem li = new ListItem();
        li.Text = "Select";
        li.Value = "0";
        drpCntry.Items.Add(li);
        drpCntry.SelectedValue = "0";
        gCloseConnection(cn);
    }
    #endregion

    #region returns Department using SP_S_dept_lkup proc.
    public void getDeptmentList(DropDownList drpDept)
    {
        string sqlstr = string.Empty;
        DataSet DS;
        cn = gOpenConnection();
        sqlstr = " sp_s_dept_lkup " ;
        DS = getDataSet(sqlstr);
        if (DS.Tables[0].Rows.Count > 0)
        {
            DataTable dt = new DataTable();
            dt = DS.Tables[0];
            drpDept.DataSource = dt;
            drpDept.DataTextField = "NameTxt";
            drpDept.DataValueField = "IdNmb";
            drpDept.DataBind();
        }
        else
            drpDept.SelectedIndex = drpDept.Items.Count - 1;
        ListItem li = new ListItem();
        li.Text = "Select";
        li.Value = "0";
        drpDept.Items.Add(li);
        drpDept.SelectedValue = "0";
        gCloseConnection(cn);
    }
    #endregion

    #region returns Claims using sp_s_form_w8ben_claim_lkup proc.
    public void getClaimList(CheckBoxList chk, string sqlstr, string NameTxt, string IdNmb)
    {
        //string sqlstr = string.Empty;
        DataSet DS;
        cn = gOpenConnection();
        //sqlstr = " SP_S_state_lkup " + CountryId;
        DS = getDataSet(sqlstr);
        if (DS.Tables[0].Rows.Count > 0)
        {
            DataTable dt = new DataTable();
            dt = DS.Tables[0];
            chk.DataSource = dt;
            chk.DataTextField = NameTxt;
            chk.DataValueField = IdNmb;
            chk.DataBind();
        }

        gCloseConnection(cn);
    }
    public void getClaimList(ListBox list, string sqlstr, string NameTxt, string IdNmb)
    {
        //string sqlstr = string.Empty;
        DataSet DS;
        cn = gOpenConnection();
        //sqlstr = " SP_S_state_lkup " + CountryId;
        DS = getDataSet(sqlstr);
        if (DS.Tables[0].Rows.Count > 0)
        {
            DataTable dt = new DataTable();
            dt = DS.Tables[0];
            list.DataSource = dt;
            list.DataTextField = NameTxt;
            list.DataValueField = IdNmb;
            list.DataBind();
        }

        gCloseConnection(cn);
    }
     #endregion
        #region DTS File Paths
    public string FileName(string dtsFileName)
    {
        string Path = string.Empty;
        if (dtsFileName.ToLower() == "funddistribution")
        {

            Path = "Distribution Sample File Layout.xlsx";
        }
        else if (dtsFileName.ToLower() == "capitalcall")
        {

            Path = "Capital Call Sample File Layout.xlsx";
        }
        else if (dtsFileName.ToLower() == "disttorecocapitalcall")
        {

            Path = "GDGSCapitalCallSampleFileLayout.xlsx";
        }
        else if (dtsFileName.ToLower() == "disttorecodistribution")
        {

            Path = "GDGSDistributionSampleFileLayout.xlsx";
        }

        return Path;
    }
    #endregion 
}
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           