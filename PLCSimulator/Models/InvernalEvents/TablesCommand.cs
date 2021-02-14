using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using PLCTools.Common;
using PLCTools.Service;

namespace PLCTools.Models.InvernalEvents
{
    public class TablesCommand
    {
        public string ColumnName { get; set; }
        public string Command { get; set; }
        public List<ColumnsCommand> Columns { get; set; }
        public string Where { get; set; }
        public void Execute(Boolean ReplaceAndExecute, string tagevent,TagHandler Tags)
        {
            string finalWhere;
            string value;
            string cmdSend;
            Logging log = new Logging("[" + tagevent + "]" + Command + ":" + ColumnName + " ");
            SQLHandler sqlGenerator = new SQLHandler();

            string cmd = Command.ToLower();
            if (cmd.Substring(0, 6) == "insert" || cmd.Substring(0, 6) == "update" || cmd.Substring(0, 6) == "select")
            {
                if (Where != "")
                {
                    finalWhere = Tags.SolveWhere(Where);
                }
                else
                {
                    finalWhere = Where;
                }

                sqlGenerator.InitCommand(cmd.ToString(), ColumnName.ToString(), finalWhere);
                if (Columns != null)
                {
                    for (int index = 0; index < Columns.Count; index++)
                    {
                        ColumnsCommand column = Columns[index];
                        log.Command += column.TagAddress + column.name + column.KeepInTag + "=" + column.Formula + "\t";
                        if (column.Enabled)
                        {
                            if (ReplaceAndExecute)
                            {
                                string tmpExp = column.Formula;
                                value = Tags.SolveFormula(tmpExp);
                                if (column.Format != "")
                                {
                                    value = String.Format(column.Formula, column.Format);
                                }
                            }
                            else
                            {
                                value = column.Formula;
                            }
                            column.Value = value.ToString().Trim();
                            sqlGenerator.SetColumn(column.name, value);

                            if (column.KeepInTag != "")
                            {
                                Tags.UpdateTagsValue(column.KeepInTag, Convert.ToInt32(value));
                                IntData.SaveRecordsetFlag = true;
                            }
                        }
                        Columns[index] = column;
                    }

                    if (ReplaceAndExecute)
                    {
                        cmdSend = sqlGenerator.Command;
                        if (Command.ToString().ToUpper() == "UPDATE")
                        {
                            if (AutoInsert(ColumnName.ToString()))
                            {
                                cmdSend += IntData.AUTO_INSERT_STR;
                            }
                        }
                        if (Command.ToString().ToUpper() != "SELECT" && ColumnName.ToString() != "" && ColumnName.IndexOf("_Internal") == -1)
                        {
                            MQHandler queue = new MQHandler();
                            queue.SendMsg(cmdSend);
                        }
                        if (Command.ToString().ToUpper() == "SELECT")
                        {
                            WriteOPC(cmdSend, Tags);
                        }
                    }
                }

            }
            log.Success();
        }
        private static Boolean AutoInsert(string TableName)
        {
            foreach (Table tt in IntData.TableList)
            {
                if (tt.Name == TableName)
                {
                    return tt.AutoInsert;
                }
            }
            return false;
        }
        private void WriteOPC(string sqlQuery,TagHandler Tags)
        {
            Logging log = new Logging(sqlQuery);
            string tableName = "";
            string operation = "";
            int CantTags = 0;
            SQLHandler.ParseTable(sqlQuery, ref tableName, ref operation);
            List<string> TagValues = new List<string>();
            List<string> TagNames = new List<string>();
            List<string> ServerNames = new List<string>();
            List<string> PLCNames = new List<string>();
            List<string> Register = new List<string>();

            if (tableName == "PLC2DB_Internal")
            {
                CantTags = Columns.Count;
                foreach (ColumnsCommand cc in Columns)
                {
                    string value = cc.Value;
                    TagValues.Add(cc.Format == "" ? value : String.Format(value, cc.Format));
                    TagNames.Add(cc.WriteToTag);
                    ServerNames.Add(cc.ServerName);
                    PLCNames.Add(cc.PLCName);
                    Register.Add(cc.TagAddress);
                }
            }
            else
            {
                try
                {
                    string ConnStr = IntData.DestConn;
                    using (var connection = new SqlConnection(ConnStr))
                    {
                        connection.Open();
                        SqlCommand cmd = new SqlCommand(sqlQuery, connection);
                        SqlDataReader sdr = cmd.ExecuteReader();
                        CantTags = sdr.FieldCount;

                        if (sdr.FieldCount > 0)
                        {
                            for (int i = 0; i < CantTags; i++)
                            {
                                sdr.Read();
                                ColumnsCommand cc = Columns[i];
                                TagValues.Add(sdr[0].ToString());
                                TagNames.Add(cc.WriteToTag.Trim() == "" ? cc.KeepInTag : cc.WriteToTag);
                                ServerNames.Add(cc.ServerName);
                                PLCNames.Add(cc.PLCName);
                                Register.Add(cc.TagAddress);
                            }
                        }
                        else if (tableName == "PI_Furn_Plan")
                        {
                            for (int i = 0; i < CantTags; i++)
                            {
                                sdr.Read();
                                ColumnsCommand cc = Columns[i];
                                TagValues.Add("0");
                                TagNames.Add(cc.WriteToTag == "" ? cc.KeepInTag : cc.WriteToTag);
                                ServerNames.Add(cc.ServerName);
                                PLCNames.Add(cc.PLCName);
                                Register.Add(cc.TagAddress);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    log.Fatal(ex);
                }
            }
            //-----Perform specific changes and calculation with the record read from database
            if (tableName.ToLower() == "ophtdata")
            {
                //If they change the heat number, also change the year (1st part of the heat number)
                int ix1 = 0;
                int ix2 = 0;
                for (int j = 0; j < TagValues.Count; j++)
                {
                    if (Columns[j].Formula == "CCM_HEAT_1")
                    {
                        ix1 = j;
                    }
                    if (Columns[j].Formula == "CCM_HEAT_2")
                    {
                        ix2 = j;
                    }
                }

                if (Columns[ix2].LastValue != Convert.ToInt32(TagValues[ix2]))
                {
                    ColumnsCommand tt = Columns[ix1];
                    tt.LastValue = -1;
                    Columns[ix1] = tt;
                }
            }
            // ----------------------------------------- Write to the PLC

            int TagIndex = 0;
            int idxPLCName;


            if (ServerNames.Count != 0)
            {
                idxPLCName = -1;
                for (int i = 0; i < ServerNames.Count; i++)
                {
                    string serNam = ServerNames[i].ToString().Trim();
                    if (serNam != "-" && serNam != "")
                    {
                        idxPLCName = i;
                        break;
                    }
                }

                if (idxPLCName != -1)
                {
                    using (OPCController OPCGrp_w = new OPCController(ServerNames[idxPLCName], "Write", PLCNames[idxPLCName]))
                    {
                        System.Array str = new object[CantTags + 1];
                        List<string> PLCTagNames = new List<string>();
                        for (int i = 0; i < CantTags; i++)
                        {
                            if (Columns[i].ServerName != "" && Columns[i].ServerName != "-")
                            {
                                TagIndex++;
                                str.SetValue(TagValues[i], i + 1);
                                PLCTagNames.Add(TagNames[i]);
                                OPCGrp_w.AddItem(TagNames[i], Register[i], " ", PLCNames[i]);
                            }
                        }
                        OPCGrp_w.PutData(str);
                    }
                }
            }
            for (int i = 0; i < CantTags; i++)
            {
                ColumnsCommand cct = Columns[i];
                if (Columns[i].ServerName != "" && Columns[i].ServerName != "-")
                {
                    cct.LastValue = Convert.ToInt32(TagValues[i]);
                    Columns[i] = cct;
                }
                else
                {
                    Tags.UpdateTagsValue(TagNames[i], TagValues[i]);
                }
            }
        }
    }
}
