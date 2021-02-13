using System;
using System.Data.SqlClient;
using PLCTools.Common;
using PLCTools.Models;

namespace PLCTools.Service
{
    class SQLHandler
    {
        const int CMD_INSERT = 1;
        const int CMD_UPDATE = 2;
        const int CMD_SELECT = 3;

        private string SqlCmdString;
        private string SqlCmd1;
        private string SqlCmd2;
        private string SqlCmdCompose1;
        private string SqlCmdCompose2;
        private int SqlCmdId;

        public Boolean InitCommand(string command, string table, string where)
        {
            if (command.ToLower() == "insert")
            {
                SqlCmdId = CMD_INSERT;
                SqlCmd1 = "";
                SqlCmd2 = "";

                SqlCmdCompose1 = "insert into " + table + " (";
                SqlCmdCompose2 = " values (";
            }

            if (command.ToLower() == "update")
            {
                SqlCmdId = CMD_UPDATE;
                SqlCmd1 = "";
                SqlCmd2 = "";

                SqlCmdCompose1 = "update " + table + " set ";
                SqlCmdCompose2 = " " + where;
            }

            if (command.ToLower() == "select")
            {
                SqlCmdId = CMD_SELECT;
                SqlCmd1 = "";
                SqlCmd2 = "";

                SqlCmdCompose1 = "select ";
                SqlCmdCompose2 = "from " + table + " " + where;
            }
            return true;
        }
        public Boolean SetColumn(string column, string value)
        {
            if (SqlCmdId == CMD_INSERT)
            {
                if (SqlCmd1.Length > 0)
                {
                    SqlCmd1 += ",";
                    SqlCmd2 += ",";
                }
                SqlCmd1 += column;
                if (value.ToUpper() == "STRING")
                {
                    if (IsDBFunction(value))
                    {
                        SqlCmd2 += value;
                    }
                    else
                    {
                        SqlCmd2 += "'" + value + "'";
                    }
                }
                else if (value.ToUpper() == "DATE")
                {
                    SqlCmd2 += "'" + value + "'";
                }
                else
                {
                    SqlCmd2 += value;
                }
            }
            else if (SqlCmdId == CMD_UPDATE)
            {
                if (SqlCmd1.Length > 0)
                {
                    SqlCmd1 += ",";
                }
                SqlCmd1 += column + "=";
                if (value.ToUpper() == "STRING")
                {
                    if (IsDBFunction(value))
                    {
                        SqlCmd1 += value;
                    }
                    else
                    {
                        SqlCmd1 += "'" + value + "'";
                    }

                }
                else if (value.ToUpper() == "DATE")
                {
                    SqlCmd1 += "'" + value + "'";
                }
                else
                {
                    SqlCmd1 += value;
                }

            }
            else if (SqlCmdId == CMD_SELECT)
            {
                if (SqlCmd1.Length > 0)
                {
                    SqlCmd1 += ",";
                    //SqlCmd2 = SqlCmd2 + ",";
                }
                SqlCmd1 += column;

                if (value.ToUpper() == "STRING")
                {
                    //SqlCmd2 = SqlCmd2 + "'" + value + "'";
                }
                else if (value.ToUpper() == "DATE")
                {
                    //SqlCmd2 = SqlCmd2 + "'" + value + "'";
                }
                else
                {
                    //SqlCmd2 = SqlCmd2 + value;
                }

            }
            return true;
        }
        public string Command
        {
            set
            {
                SqlCmdCompose1 = "";
                SqlCmdCompose2 = "";
                SqlCmdString = value;
            }
            get
            {
                if (SqlCmdCompose1 != "")
                {
                    if (SqlCmdId == CMD_INSERT)
                    {
                        SqlCmdString = SqlCmdCompose1 + SqlCmd1 + ")";
                        SqlCmdString = SqlCmdString + SqlCmdCompose2 + SqlCmd2 + ")";
                    }
                    if (SqlCmdId == CMD_UPDATE)
                    {
                        SqlCmdString = SqlCmdCompose1 + SqlCmd1 + SqlCmdCompose2;
                    }
                    if (SqlCmdId == CMD_SELECT)
                    {
                        SqlCmdString = SqlCmdCompose1 + SqlCmd1 + " ";
                        SqlCmdString = SqlCmdString + SqlCmdCompose2 + SqlCmd2;
                    }
                }
                return SqlCmdString;
            }
        }
        private Boolean IsDBFunction(string command)
        {
            switch (command)
            {
                case "getdate()": return true;
                default: return false;
            }
        }
        public Boolean ExecuteCC()
        {
            Logging log = new Logging(SqlCmdString);
            try
            {
                using (SqlConnection Connection = new SqlConnection(IntData.CfgConn))
                {
                    Connection.Open();
                    SqlCommand SqlCmd = new SqlCommand
                    {
                        CommandText = SqlCmdString,
                        Connection = Connection
                    };
                    SqlCmd.ExecuteNonQuery();
                }
                log.Success();
                return true;
            }
            catch (Exception ex)
            {
                log.Fatal(ex);
                return false;
            }
        }
        public static string GenerateInsert(string commandStr)
        {
            string strTmp;
            int posTmp;
            string tableName = "";
            string operation = "";

            strTmp = commandStr;
            posTmp = commandStr.IndexOf("update ");
            string column = "";
            string value = "";
            if (posTmp > 0)
            {
                ParseTable(commandStr, ref tableName, ref operation);
                posTmp = strTmp.IndexOf(" where ");

                strTmp = strTmp.Substring("where ".Length + posTmp);
                posTmp = strTmp.IndexOf("=");
                column = strTmp.Substring(0, posTmp - 1);
                value = strTmp.Substring(strTmp.Length - posTmp);
            }

            return "insert into " + tableName + "(" + column + ") values(" + value + ")";
        }
        public static void ParseTable(string command, ref string TableName, ref string SQLOperation)
        {
            string subStr = "";
            int PosS = command.ToUpper().IndexOf("SELECT ");
            int PosF = command.ToUpper().IndexOf("FROM ");
            int PosI = command.ToUpper().IndexOf("INSERT INTO");
            int PosU = command.ToUpper().IndexOf("UPDATE");


            if (PosI + 1 > 0)
            {
                subStr = command.Substring(PosI + 11).Trim();
                SQLOperation = "INSERT";
            }
            else if (PosU + 1 > 0)
            {
                subStr = command.Substring(PosU + 7).Trim();
                SQLOperation = "UPDATE";
            }
            else if (PosS + PosF + 2 > 0)
            {
                subStr = command.Substring(PosF + 4).Trim();
                SQLOperation = "SELECT";
            }

            int Pos = 0;
            while (subStr.Substring(Pos, 1) != " " && Pos < subStr.Length)
            {
                TableName = String.Concat(TableName, subStr.Substring(Pos++, 1));
            }

        }
    }
}
