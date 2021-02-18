using System;
using System.Collections;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;

namespace PLCTools.Service
{
    /// <summary>
    /// Summary description for DbCommon.
    /// </summary>
    public class DbCommon : BaseClass
    {
        public string OleConnectionStringSource { get; set; }
        public string OleConnectionStringTarget { get; set; }
        public DbCommon()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        public OleDbConnection GetOleConnection(string taget_source)
        {
            OleDbConnection dbconnection = null;
            if (dbconnection == null)
            {
                string ls_connection;
                if (taget_source.ToUpper() == "SOURCE")
                {
                    ls_connection = OleConnectionStringSource;
                }
                else
                {
                    ls_connection = OleConnectionStringTarget;
                }
                dbconnection = new OleDbConnection(ls_connection);
            }
            actionLog("DbCommon:GetOleConnection:", "Database Connection: " + dbconnection.ConnectionString);
            return dbconnection;
        }
        public DataTable getOleProData(string ps_sql_name, OleDbConnection myConnection)
        {
            OleDbCommand cmd = new OleDbCommand(ps_sql_name, myConnection);
            DataTable dt = new DataTable();
            OleDbDataAdapter adpt = new OleDbDataAdapter(cmd);
            try
            {
                cmd.Connection = myConnection;
                cmd.CommandText = ps_sql_name;
                cmd.CommandType = CommandType.StoredProcedure;
                adpt.Fill(dt);
            }
            catch (Exception ex)
            {
                StackTrace st = new StackTrace(new StackFrame(true));
                actionLog("DbCommon:getOleProData(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", ex.ToString());
                base.sendEmail("DspBatch dbCommon:getOleProData Error", ex.ToString());
            }
            finally
            {
                if (cmd != null)
                {
                    cmd.Dispose();
                }
                if (adpt != null)
                {
                    adpt.Dispose();
                }
            }
            return dt;
        }
        public DataTable getOleProData(string ps_sql_name, ArrayList ps_parm, OleDbConnection myConnection)
        {
            OleDbCommand cmd = new OleDbCommand(ps_sql_name, myConnection);
            DataSet dset;
            dset = new DataSet();
            DataTable dt;
            dt = new DataTable();
            OleDbDataAdapter adpt = new OleDbDataAdapter(cmd);
            try
            {
                cmd.Connection = myConnection;
                actionLog("DbCommon:getOleProData:", myConnection.DataSource.ToString() + "." + ps_sql_name);
                cmd.CommandText = ps_sql_name;
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                for (int i = 0; i < ps_parm.Count; i++)
                {
                    cmd.Parameters.Add(ps_parm[i]);
                    actionLog("DbCommon:getOleProData:", "Parameters: " + cmd.Parameters[i].ParameterName.ToString() + ", " + cmd.Parameters[i].Size.ToString() + ", " + cmd.Parameters[i].Value.ToString());
                }
                actionLog("DbCommon:getOleProData:", "Calling adpt.Fill Function.");
                adpt.Fill(dset);
                dt = dset.Tables[0];
                actionLog("DbCommon:getOleProData:", "adpt.Fill Function Returned " + dt.Rows.Count.ToString() + " row(s)");
            }
            catch (Exception ex)
            {
                StackTrace st = new StackTrace(new StackFrame(true));
                actionLog("DbCommon:getOleProData(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Execution Error: " + ex.ToString());
                base.sendEmail("DspBatch:DbCommon:getOleProData Error", ex.ToString());
            }
            finally
            {
                if (cmd != null)
                {
                    cmd.Dispose();
                }
                if (adpt != null)
                {
                    adpt.Dispose();
                }
            }
            return dt;
        }
        public DataTable getOleTableData(string ps_sql_name, OleDbConnection myConnection)
        {
            OleDbCommand cmd = new OleDbCommand(ps_sql_name, myConnection);
            DataTable dt = new DataTable();
            OleDbDataAdapter adpt = new OleDbDataAdapter(cmd);
            try
            {
                cmd.Connection = myConnection;
                cmd.CommandText = ps_sql_name;
                cmd.CommandType = CommandType.Text;
                adpt.Fill(dt);
            }
            catch (Exception ex)
            {
                StackTrace st = new StackTrace(new StackFrame(true));
                actionLog("DbCommon:getOleTableData(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", ex.ToString());
                base.sendEmail("DspBatch:DbCommon:getOleTableData Error", ex.ToString());
            }
            finally
            {
                if (cmd != null)
                {
                    cmd.Dispose();
                }
                if (adpt != null)
                {
                    adpt.Dispose();
                }
            }
            return dt;
        }
        public DataTable getOleTableData(string ps_sql_name, ArrayList ps_parm, OleDbConnection myConnection)
        {
            OleDbCommand cmd = new OleDbCommand(ps_sql_name, myConnection);
            DataTable dt = new DataTable();
            OleDbDataAdapter adpt = new OleDbDataAdapter(cmd);
            try
            {
                cmd.Connection = myConnection;
                cmd.CommandText = ps_sql_name;
                cmd.CommandType = CommandType.Text;
                for (int i = 0; i < ps_parm.Count; i++)
                {
                    cmd.Parameters.Add(ps_parm[i]);
                }
                adpt.Fill(dt);
            }
            catch (Exception ex)
            {
                StackTrace st = new StackTrace(new StackFrame(true));
                actionLog("DbCommon:getOleTableData(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", ex.ToString());
                base.sendEmail("DspBatch:DbCommon:getOleTableData Error", ex.ToString());
            }
            finally
            {
                if (cmd != null)
                {
                    cmd.Dispose();
                }
                if (adpt != null)
                {
                    adpt.Dispose();
                }
            }
            return dt;
        }
        public int updOleProData(string ps_sql_name, OleDbConnection myConnection, OleDbTransaction myTrans)
        {
            int li_return = 1;
            OleDbCommand cmd = new OleDbCommand(ps_sql_name, myConnection);
            try
            {
                cmd.Connection = myConnection;
                cmd.Transaction = myTrans;
                cmd.CommandText = ps_sql_name;
                cmd.CommandType = CommandType.StoredProcedure;
                li_return = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                li_return = -1;
                StackTrace st = new StackTrace(new StackFrame(true));
                actionLog("DbCommon:updOleProData(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", ex.ToString());
                base.sendEmail("DspBatch:DbCommon:updOleProData Error", ex.ToString());
            }
            finally
            {
                if (cmd != null)
                {
                    cmd.Dispose();
                }
            }
            return li_return;
        }
        public int updOleProData(string ps_sql_name, ArrayList ps_parm, OleDbConnection myConnection, OleDbTransaction myTrans)
        {
            int li_return = 1;
            OleDbCommand cmd = new OleDbCommand(ps_sql_name, myConnection);
            try
            {
                cmd.Connection = myConnection;
                cmd.Transaction = myTrans;
                cmd.CommandText = ps_sql_name;
                cmd.CommandType = CommandType.StoredProcedure;
                for (int i = 0; i < ps_parm.Count; i++)
                {
                    cmd.Parameters.Add(ps_parm[i]);
                }
                li_return = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                li_return = -1;
                StackTrace st = new StackTrace(new StackFrame(true));
                actionLog("DbCommon:updOleProData(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", ex.ToString());
                base.sendEmail("DspBatch:DbCommon:updOleProData Error", ex.ToString());
            }
            finally
            {
                if (cmd != null)
                {
                    cmd.Dispose();
                }
            }
            return li_return;
        }
        public int updOleProData(string ps_sql_name, OleDbConnection myConnection)
        {
            int li_return = 1;
            OleDbCommand cmd = new OleDbCommand(ps_sql_name, myConnection);
            try
            {
                cmd.Connection = myConnection;
                cmd.CommandText = ps_sql_name;
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                li_return = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                li_return = -1;
                StackTrace st = new StackTrace(new StackFrame(true));
                actionLog("DbCommon:updOleProData(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", ex.ToString());
                base.sendEmail("DspBatch:DbCommon:updOleProData Error", ex.ToString());
            }
            finally
            {
                if (cmd != null)
                {
                    cmd.Dispose();
                }
            }
            return li_return;
        }
        public int updOleProData(string ps_sql_name, ArrayList ps_parm, OleDbConnection myConnection)
        {
            int li_return = 1;
            OleDbCommand cmd = new OleDbCommand(ps_sql_name, myConnection);
            try
            {
                cmd.Connection = myConnection;
                cmd.CommandText = ps_sql_name;
                cmd.CommandType = CommandType.StoredProcedure;
                for (int i = 0; i < ps_parm.Count; i++)
                {
                    cmd.Parameters.Add(ps_parm[i]);
                }
                li_return = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                li_return = -1;
                StackTrace st = new StackTrace(new StackFrame(true));
                actionLog("DbCommon:updOleProData(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", ex.ToString());
                base.sendEmail("DspBatch:DbCommon:updOleProData Error", ex.ToString());
            }
            finally
            {
                if (cmd != null)
                {
                    cmd.Dispose();
                }
            }
            return li_return;
        }
        public int updOleTableData(string ps_sql_name, OleDbConnection myConnection, OleDbTransaction myTrans)
        {
            int li_return = 1;
            OleDbCommand cmd = new OleDbCommand(ps_sql_name, myConnection);
            try
            {
                cmd.Connection = myConnection;
                cmd.Transaction = myTrans;
                cmd.CommandText = ps_sql_name;
                cmd.CommandType = CommandType.Text;
                li_return = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                li_return = -1;
                StackTrace st = new StackTrace(new StackFrame(true));
                actionLog("dbCommon:updOleTableData(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", ex.ToString());
                base.sendEmail("DspBatch:dbCommon:updOleTableData Error", ex.ToString());
            }
            finally
            {
                if (cmd != null)
                {
                    cmd.Dispose();
                }
            }
            return li_return;
        }
        public int updOleTableData(string ps_sql_name, ArrayList ps_parm, OleDbConnection myConnection, OleDbTransaction myTrans)
        {
            int li_return = 1;
            OleDbCommand cmd = new OleDbCommand(ps_sql_name, myConnection);
            try
            {
                cmd.Connection = myConnection;
                cmd.Transaction = myTrans;
                cmd.CommandText = ps_sql_name;
                cmd.CommandType = System.Data.CommandType.Text;
                for (int i = 0; i < ps_parm.Count; i++)
                {
                    cmd.Parameters.Add(ps_parm[i]);
                }
                li_return = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                li_return = -1;
                StackTrace st = new StackTrace(new StackFrame(true));
                actionLog("dbCommon:updOleTableData(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", ex.ToString());
                base.sendEmail("DspBatch:dbCommon:updOleTableData Error", ex.ToString());
            }
            finally
            {
                if (cmd != null)
                {
                    cmd.Dispose();
                }
            }
            return li_return;
        }
        //make in and out parameters
        public ArrayList MakeInParm(ArrayList as_name, ArrayList as_type, ArrayList as_size, ArrayList ao_value)
        {
            ArrayList parmList = new ArrayList();
            for (int i = 0; i < as_name.Count; i++)
            {
                OleDbParameter parm = new OleDbParameter();
                try
                {
                    switch ((string)as_type[i].ToString().ToLower())
                    {
                        case "char":
                            parm.OleDbType = OleDbType.Char;
                            break;
                        case "varchar":
                            parm.OleDbType = OleDbType.VarChar;
                            break;
                        case "integer":
                            parm.OleDbType = OleDbType.Integer;
                            break;
                        case "int":
                            parm.OleDbType = OleDbType.Integer;
                            break;
                        case "decimal":
                            parm.OleDbType = OleDbType.Decimal;
                            break;
                        case "datetime":
                            parm.OleDbType = OleDbType.Date;
                            break;
                        case "date":
                            parm.OleDbType = OleDbType.Date;
                            break;
                        default:
                            break;
                    }
                    parm.ParameterName = as_name[i].ToString();
                    parm.Size = int.Parse(as_size[i].ToString());
                    parm.Direction = System.Data.ParameterDirection.Input;
                    parm.Value = ao_value[i];
                    parmList.Add(parm);
                    actionLog("DbCommon:MakeInParm:", "Parm Name: " + as_name[i].ToString() + " Size: " + as_size[i].ToString() + " Value: " + ao_value[i].ToString());
                }
                catch (Exception ex)
                {
                    StackTrace st = new StackTrace(new StackFrame(true));
                    actionLog("DbCommon:MakeInParm(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", ex.ToString());
                    base.sendEmail("DspBatch:DbCommon:MakeInParm Error", ex.ToString());
                }
                finally
                {
                    if (parm != null)
                    {
                        parm = null;
                    }
                }
            }
            return parmList;
        }
        public ArrayList MakeOutParm(ArrayList as_name, ArrayList as_type, ArrayList as_size, ArrayList ao_value)
        {
            ArrayList parmList = new ArrayList();
            for (int i = 0; i < as_name.Count; i++)
            {
                OleDbParameter parm = new OleDbParameter();
                try
                {
                    switch ((string)as_type[i].ToString().ToLower())
                    {
                        case "char":
                            parm.OleDbType = OleDbType.Char;
                            break;
                        case "varchar":
                            parm.OleDbType = OleDbType.VarChar;
                            break;
                        case "integer":
                            parm.OleDbType = OleDbType.Integer;
                            break;
                        case "int":
                            parm.OleDbType = OleDbType.Integer;
                            break;
                        case "decimal":
                            parm.OleDbType = OleDbType.Decimal;
                            break;
                        case "datetime":
                            parm.OleDbType = OleDbType.Date;
                            break;
                        case "date":
                            parm.OleDbType = OleDbType.Date;
                            break;
                        default:
                            break;
                    }
                    parm.ParameterName = as_name[i].ToString();
                    parm.Size = int.Parse(as_size[i].ToString());
                    parm.Direction = System.Data.ParameterDirection.Output;
                    parm.Value = ao_value[i];
                    parmList.Add(parm);
                }
                catch (Exception ex)
                {
                    StackTrace st = new StackTrace(new StackFrame(true));
                    actionLog("DbCommon:MakeOutParm(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", ex.ToString());
                    base.sendEmail("DspBatch:DbCommon:MakeOutParm Error", ex.ToString());
                }
                finally
                {
                    if (parm != null)
                    {
                        parm = null;
                    }
                }
            }
            return parmList;
        }
    }
}
