using PLCTools.Common;
using PLCTools.Service.MathPaser;
using PLCTools.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace PLCTools.Service
{
    public class TagHandler
    {
        public List<InternalTag> Values = new List<InternalTag>();
        public void Init()
        {
            Logging log = new Logging("select Tag, OPCName, Type from PLC2DB_Tags order by tag");
            try
            {
                using (var connection = new SqlConnection(IntData.CfgConn))
                {
                    connection.Open();
                    SqlCommand cmd = new SqlCommand(log.Command, connection);
                    SqlDataReader sdr = cmd.ExecuteReader();
                    while (sdr.Read())
                    {
                        InternalTag dt = new InternalTag
                        {
                            Tag = sdr[0].ToString().Trim(),
                            IsIntTag = sdr[1].ToString().Trim() == "-" && sdr[2].ToString().Trim() == "VAR",
                            Activated = false,
                            Deactivated = false,
                            LastValue = 0,
                            Value = 0
                        };
                        Values.Add(dt);
                    }
                }
                log.Success();
            }
            catch (Exception ex)
            {
                log.Fatal(ex);
            }
        }
        public object GetTagsValue(string tag)
        {
            if (tag == "1") return 1;
            foreach (InternalTag dt in Values)
            {
                if (dt.Tag.Trim() == tag.Trim())
                {
                    return dt.Value;
                }
            }
            return ErrorCode("TAG_NOT_FOUND");
        }
        public bool UpdateTagsValue(string tag, object newValue)
        {
            for (int i = 0; i < Values.Count; i++)
            {
                InternalTag dt = Values[i];
                if (dt.Tag.Trim() == tag.Trim())
                {
                    dt.Value = newValue;
                    Values[i] = dt;
                    return true;
                }
            }
            return false;
        }
        public int GetTagsInString(string context, List<string> tags)
        {
            int Position;
            string tag;
            int index;

            Position = context.IndexOf(":", 0);
            index = 0;
            while (Position > 0)
            {
                tag = context.Substring(Position, context.IndexOf(":", Position, context.Length) - Position);
                context = context.Replace(tag, "");
                index++;
                tags.Add(tag.Substring(1, tag.Length - 3));
                Position = context.IndexOf(":", 0);
            }
            return index;
        }
        public Boolean ExecuteScripts(List<List<string>> ProcedureNames, int flanc)
        {
            Logging log = new Logging();
            try
            {
                if (ProcedureNames == null) return false;
                UserProcedure up = new UserProcedure();
                if (ProcedureNames[flanc].Count != 0)
                {
                    foreach (string userProc in ProcedureNames[flanc])
                    {
                        if (userProc.Trim() != "")
                        {
                            up.ExecuteProcedure(userProc, this);
                        }
                    }
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
        private int ErrorCode(string s)
        {
            switch (s)
            {
                case "TAG_NOT_FOUND": return 0;
            }
            return 0;
        }
        /*  This function fills the TagsInFormula table:
        - delete the table TagsInFormula (clean the table)
        - Selects ALL the formulas (or tags) in the Cfg Table
        - For each formula in Cfg get the tags used (call function) and store
         in tagList() array
        - Call to the function that INSERTS all the tags used in formula
         (and for EACH formula) to the TagsInFormula table*/
        private Boolean GenerateTagsInFormulaTable()
        {
            string sqlQuery;
            List<string> tagList = new List<string>();

            sqlQuery = "delete from PLC2DB_TagsInFormula";
            Logging log = new Logging(sqlQuery);
            try
            {
                using (SqlConnection connection = new SqlConnection(IntData.CfgConn))
                {
                    connection.Open();
                    SqlCommand cmd = new SqlCommand(sqlQuery, connection);
                    cmd.ExecuteNonQuery();
                    sqlQuery = "select ID, Expression from PLC2DB_Cfg";
                    cmd = new SqlCommand(sqlQuery, connection);
                    SqlDataReader sdr = cmd.ExecuteReader();
                    while (sdr.Read())
                    {
                        int id = Convert.ToInt32(sdr[0]);
                        string Formula = sdr[1].ToString();
                        GetTagListFromFormula(Formula, ref tagList);
                        InsertTagInFormula(id, ref tagList);
                    }
                    sdr.Close();
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
        /* Arm and execute the SQL command to insert each tag of tags() array
           in TagsInFormula table*/
        private void InsertTagInFormula(int configId, ref List<string> tags)
        {
            string sqlQuery = "";
            Logging log = new Logging();
            try
            {
                using (SqlConnection connection = new SqlConnection(IntData.CfgConn))
                {
                    connection.Open();
                    SqlCommand cmd = new SqlCommand(sqlQuery, connection);
                    foreach (string tag in tags)
                    {
                        sqlQuery = "insert into PLC2DB_TagsInFormula values (" + configId.ToString() + ", '" + tag + "')";
                        log.Command = sqlQuery;
                        cmd = new SqlCommand(sqlQuery, connection);
                        cmd.ExecuteNonQuery();
                        log.Success();
                    }
                }
            }
            catch (Exception ex)
            {
                log.Fatal(ex);
            }
        }
        private Boolean GenerateTagsInWhereTable(TagHandler Tags)
        {
            string sqlQuery;
            List<string> tags = new List<string>();

            sqlQuery = "delete  from PLC2DB_TagsInWhere";
            Logging log = new Logging(sqlQuery);
            try
            {
                using (SqlConnection connection = new SqlConnection(IntData.CfgConn))
                {
                    connection.Open();
                    SqlCommand cmd = new SqlCommand(sqlQuery, connection);
                    cmd.ExecuteNonQuery();
                    sqlQuery = "select ID, Cond from PLC2DB_Cfg where Cond is Not Null and Cond <>''";
                    cmd = new SqlCommand(sqlQuery, connection);
                    SqlDataReader sdr = cmd.ExecuteReader();
                    while (sdr.Read())
                    {
                        if (Tags.GetTagsInString(sdr[1].ToString(), tags) != 0)
                        {
                            InsertTagInWhere(Convert.ToInt32(sdr[0]), ref tags);
                        }
                    }
                    sdr.Close();
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
        /* Arm and execute the SQL command to insert each tag of tags() array
           in TagsInWhere table*/
        private Boolean InsertTagInWhere(int configId, ref List<string> tags)
        {
            SQLHandler dbCommand = new SQLHandler();
            foreach (string tag in tags)
            {
                dbCommand.Command = "insert into PLC2DB_TagsInWhere values (" + configId + ", '" + tag + "')";
                dbCommand.ExecuteCC();
            }
            return true;
        }
        public int SearchTag(string tag)
        {
            for (int i = 0; i < Values.Count; i++)
            {
                if (Values[i].Tag.Trim() == tag)
                {
                    return i;
                }
            }
            return -1;
        }
        public void RefreshFromEvents(OPCGroups opcgroup)
        {
            for (int opcIndex = 1; opcIndex < opcgroup.TotalItemNumber + 1; opcIndex++)
            {
                string tag = opcgroup.GetTagName(opcIndex);
                int index = SearchTag(tag);
                if (index > -1)
                {
                    InternalTag dt = Values[index];
                    if (Convert.ToInt32(dt.LastValue) != Convert.ToInt32(dt.Value))
                    {
                        dt.LastValue = dt.Value;
                        dt.Activated = Convert.ToInt32(dt.Value) > 0;
                        dt.Deactivated = !dt.Activated;
                    }
                    else
                    {
                        if (dt.Activated)
                        {
                            dt.LastValue = 0;
                            dt.Value = 0;
                            dt.Activated = false;
                            dt.Deactivated = true;
                        }
                        else
                        {
                            dt.Deactivated = false;
                        }
                    }
                    Values[index] = dt;
                }
            }
        }
        public void RefreshFromGroup(OPCGroups opcgroup)
        {
            for (int index = 1; index < opcgroup.TotalItemNumber + 1; index++)
            {
                for (int i = 0; i < Values.Count; i++)
                {
                    InternalTag dt = Values[i];
                    if (dt.Tag == opcgroup.GetTagName(index))
                    {
                        dt.Value = opcgroup.GetTagValue(dt.Tag);
                        dt.Activated = opcgroup.Activated(dt.Tag);
                        dt.Deactivated = opcgroup.Deactivated(dt.Tag);
                    }
                    Values[i] = dt;
                }
            }
        }
        public string SolveFormula(string expression)
        {
            if (expression == "") return "";
            if (expression.Substring(0, 1) == "'")
            {
                return expression.Replace("'", "");
            }
            RPN rpn = new RPN();
            if (rpn.Parse(expression))
            {
                Stack<object> operandProcessed = new Stack<object>();
                object[] tokens = rpn.Tokens.ToArray();
                for (int i = tokens.Length - 1; i >= 0; i--)
                {
                    object item = (object)tokens[i];
                    if (item is Operator @operator)
                    {
                        operandProcessed.Push(@operator);
                    }
                    else
                    {
                        Operand operand = (Operand)item;
                        if (operand.Type == OperandType.STRING)
                        {
                            object tag = operand.Value;
                            if (IsSystemVariable((string)tag))
                            {
                                tag = CalculateSystemVariable((string)tag);
                            }
                            else
                            {
                                tag = GetTagsValue((string)tag);
                            }
                            operand.Value = tag;
                            operand.Type = OperandType.NUMBER;
                        }
                        operandProcessed.Push((object)operand);
                    }
                }
                rpn.Tokens = operandProcessed;
            }
            return rpn.Evaluate().ToString();
        }
        public string SolveWhere(string whereClause)
        {
            object value;
            string newWhere;
            int cant;

            newWhere = whereClause;
            List<string> tags = new List<string>();
            cant = GetTagsInString(whereClause, tags);
            if (cant == 0)
            {
                return whereClause;
            }

            foreach (string tag in tags)
            {
                value = GetTagsValue(tag);
                string rep = ":" + tag + ":";
                newWhere = newWhere.Replace(rep, value.ToString());
            }
            return newWhere;
        }
        public Boolean GetTagListFromFormula(string Formula, ref List<string> tags)
        {
            int i = 0;
            RPN func = new RPN();
            if (func.Parse(Formula))
            {
                foreach (object item in func.Tokens)
                {
                    if (item is Operand operand)
                    {
                        tags[i] = operand.Value.ToString();
                    }
                }
            }
            else
            {
                return false;
            }
            return true;
        }
        private Boolean IsSystemVariable(string tag)
        {
            if (tag.Substring(tag.Length - 1, 1) == "$")
            {
                return true;
            }
            return false;
        }
        private string CalculateSystemVariable(string SysVar)
        {
            if (SysVar == "TIME$")
            {
                return DateTime.Now.ToString();
            }
            else if (SysVar == "MSDBTIME$")
            {
                return "getdate()";
            }
            else
            {
                return "";
            }

        }
    }
}
