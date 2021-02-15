using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using PLCTools.Models;
using PLCTools.Common;
using PLCTools.Models.InvernalEvents;

namespace PLCTools.Service
{
    public class OPCGroup_t
    {
        public OPCController opcgroup;
        public string tagEvent;
    }
    public class EventsHandler
    {
        private const int POSITIVE_FLANC = 0;
        private const int NEGATIVE_FLANC = 1;
        private List<TagEvent> OPCEvents { get; set; } = new List<TagEvent>();
        private List<TimerEvent> TimerEvents { get; set; } = new List<TimerEvent>();
        private List<OPCGroup_t> EvGroups { get; set; } = new List<OPCGroup_t>();
        private List<OPCGroup_t> VarGroups { get; set; } = new List<OPCGroup_t>();
        private List<OPCGroup_t> WhereGroups { get; set; } = new List<OPCGroup_t>();
        public void Init()
        {
            OPCEvents = IntiOPCEvents();
            TimerEvents = InitTimerEvents();
            EvGroups = InitEventsGroup();
            VarGroups = InitVarGroup();
            WhereGroups = InitWhereGroup();

        }
        public void Refresh(TagHandler Tags)
        {

            RefreshEventsGroup(Tags);
            RefreshTimerEvents();
            IntData.SaveRecordsetFlag = false;
            for (int i = 0; i < OPCEvents.Count; i++)
            {
                if (OPCEvents[i].Enable)
                {
                    if (OPCEvents[i].Activated)
                    {
                        RefreshVarGroup(OPCEvents[i].Tagevent, Tags);
                        RefreshWhereGroup(Tags);
                        for (int index = 0; index < OPCEvents[i].Targets[POSITIVE_FLANC].Count; index++)
                        {
                            OPCEvents[i].Targets[POSITIVE_FLANC][index].Execute(true, OPCEvents[i].Tagevent, Tags);
                        }
                        Tags.ExecuteScripts(OPCEvents[i].ScriptsList, POSITIVE_FLANC);
                    }
                    if (OPCEvents[i].Deactivated)
                    {
                        RefreshVarGroup(OPCEvents[i].Tagevent, Tags);
                        RefreshWhereGroup(Tags);
                        for (int index = 0; index < OPCEvents[i].Targets[NEGATIVE_FLANC].Count; index++)
                        {
                            OPCEvents[i].Targets[NEGATIVE_FLANC][index].Execute(true, OPCEvents[i].Tagevent, Tags);
                        }
                        Tags.ExecuteScripts(OPCEvents[i].ScriptsList, NEGATIVE_FLANC);
                    }
                }
            }
            if (IntData.SaveRecordsetFlag)
            {
                //***JB1 MsgBox "grabarRecordsetFlag"
                //tcklog.log "Saving Record Set"
            }
            for (int i = 0; i < TimerEvents.Count; i++)
            {
                if (TimerEvents[i].enable)
                {
                    if (TimerEvents[i].Activated)
                    {
                        RefreshVarGroup(TimerEvents[i].Tagevent, Tags);
                        RefreshWhereGroup(Tags);
                        for (int index = 0; index < TimerEvents[i].Targets[POSITIVE_FLANC].Count; index++)
                        {
                            TimerEvents[i].Targets[POSITIVE_FLANC][index].Execute(true, TimerEvents[i].Tagevent, Tags);
                        }
                        Tags.ExecuteScripts(TimerEvents[i].ScriptsList, POSITIVE_FLANC);
                    }
                    if (TimerEvents[i].Deactivated)
                    {
                        RefreshVarGroup(TimerEvents[i].Tagevent, Tags);
                        RefreshWhereGroup(Tags);
                        for (int index = 0; index < TimerEvents[i].Targets[NEGATIVE_FLANC].Count; index++)
                        {
                            TimerEvents[i].Targets[NEGATIVE_FLANC][index].Execute(true, TimerEvents[i].Tagevent, Tags);
                        }
                        Tags.ExecuteScripts(TimerEvents[i].ScriptsList, NEGATIVE_FLANC);
                    }
                }
            }
        }
        private List<TagEvent> IntiOPCEvents()
        {
            List<TagEvent> Events = new List<TagEvent>();
            string sqlStr = "select TagEvent, PLC2DB_Tags.Enabled as Enabled from PLC2DB_Cfg, PLC2DB_Tags where right(rtrim(TagEvent),1) <> '$' and PLC2DB_Tags.Enabled  is null or PLC2DB_Tags.Enabled = 1 and TagEvent = PLC2DB_Tags.Tag group by TagEvent,  PLC2DB_Tags.Enabled";
            Logging log = new Logging(sqlStr);
            try
            {
                using (SqlConnection connection = new SqlConnection(IntData.CfgConn))
                {
                    connection.Open();
                    SqlCommand cmd = new SqlCommand(sqlStr, connection);
                    SqlDataReader sdr = cmd.ExecuteReader();
                    while (sdr.Read())
                    {
                        TagEvent ev = new TagEvent
                        {
                            Tagevent = sdr[0].ToString().Trim(),
                            Enable = Convert.ToBoolean(sdr[1])
                        };
                        ev.Targets = MakeCommandsList(ev.Targets, ev.Tagevent);
                        ev.ScriptsList = FilluserProcedures(ev.ScriptsList, ev.Tagevent);
                        Events.Add(ev);
                    }
                }
                log.Success();
            }
            catch (Exception ex)
            {
                log.Fatal(ex);
            }
            return Events;
        }
        private List<TimerEvent> InitTimerEvents()
        {
            List<TimerEvent> TimerEvents = new List<TimerEvent>();
            string sqlStr = "select Period , Unit, TagEvent, PLC2DB_Timers.Enabled as Enabled from PLC2DB_Cfg, PLC2DB_Timers where right(rtrim(TagEvent),1) = '$' and PLC2DB_Cfg.TagEvent = ltrim(rtrim(PLC2DB_Timers.Name)) group by Period , Unit, TagEvent, PLC2DB_Timers.Enabled ";
            Logging log = new Logging(sqlStr);
            try
            {
                using (SqlConnection connection = new SqlConnection(IntData.CfgConn))
                {
                    connection.Open();
                    SqlCommand cmd = new SqlCommand(sqlStr, connection);
                    SqlDataReader sdr = cmd.ExecuteReader();
                    while (sdr.Read())
                    {
                        TimerEvent ev = new TimerEvent
                        {
                            Tagevent = sdr[2].ToString().Trim(),
                            enable = (bool)sdr[3],
                            TimerPeriod = GetPeriod(Convert.ToInt32(sdr[0].ToString().Trim()), sdr[1].ToString().Trim())
                        };
                        ev.Targets = MakeCommandsList(ev.Targets, ev.Tagevent);
                        ev.ScriptsList = FilluserProcedures(ev.ScriptsList, ev.Tagevent);
                        TimerEvents.Add(ev);
                    }
                    sdr.Close();
                }
                log.Success();
            }
            catch (Exception ex)
            {
                log.Fatal(ex);
            }
            return TimerEvents;
        }
        private List<OPCGroup_t> InitEventsGroup()
        {
            string sqlQuery;
            sqlQuery = "select OpcName from PLC2DB_Tags where type = 'EV' group by OpcName";
            Logging log = new Logging(sqlQuery);
            List<OPCGroup_t> opt = new List<OPCGroup_t>();
            try
            {
                using (SqlConnection connection = new SqlConnection(IntData.CfgConn))
                {
                    connection.Open();
                    SqlCommand cmd = new SqlCommand(sqlQuery, connection);
                    SqlDataReader sdr = cmd.ExecuteReader();
                    while (sdr.Read())
                    {
                        OPCGroup_t og = new OPCGroup_t
                        {
                            tagEvent = "",
                            opcgroup = new OPCController()
                        };
                        og.opcgroup.ServerName = sdr[0].ToString().Trim();
                        og.opcgroup.GroupName = "evGroup" + opt.Count;
                        opt.Add(og);
                        sqlQuery = "select PLCName, Tag, Register, Description from PLC2DB_Tags where type = 'EV' and OPCName='" + og.opcgroup.ServerName + "'";
                        CreateItemsFromSelect(sqlQuery, ref og.opcgroup);
                    }
                    sdr.Close();
                }
                log.Success();
                return opt;
            }
            catch (Exception ex)
            {
                log.Fatal(ex);
                return null;
            }
        }
        private List<OPCGroup_t> InitVarGroup()
        {
            string sqlQuery;
            sqlQuery = "select TagEvent, OpcName from PLC2DB_Tags, PLC2DB_Cfg, PLC2DB_TagsInFormula where type = 'VAR' and PLC2DB_Tags.Tag = PLC2DB_TagsInFormula.tag and PLC2DB_TagsInFormula.CfgID = PLC2DB_CFg.ID and len(rtrim(OpcName))>1 group by OpcName, TagEvent";
            Logging log = new Logging(sqlQuery);
            List<OPCGroup_t> varGroups = new List<OPCGroup_t>();
            try
            {
                using (SqlConnection connection = new SqlConnection(IntData.CfgConn))
                {
                    connection.Open();
                    SqlCommand cmd = new SqlCommand(sqlQuery, connection);
                    SqlDataReader sdr = cmd.ExecuteReader();
                    while (sdr.Read())
                    {
                        OPCGroup_t og = new OPCGroup_t
                        {
                            tagEvent = sdr[0].ToString(),
                            opcgroup = new OPCController()
                        };
                        og.opcgroup.ServerName = sdr[1].ToString().Trim();
                        og.opcgroup.GroupName = "varGroup" + varGroups.Count;

                        sqlQuery = "select PLC2DB_Tags.Tag, Register, Description, PLCName from PLC2DB_Tags, PLC2DB_Cfg, PLC2DB_TagsInFormula where type = 'VAR' and OPCName  ='" + og.opcgroup.ServerName + "' and PLC2DB_Tags.Tag = PLC2DB_TagsInFormula.tag and PLC2DB_TagsInFormula.CfgID = PLC2DB_CFg.ID and PLC2DB_Cfg.TagEvent= '" + og.tagEvent + "'";
                        CreateItemsFromSelect(sqlQuery, ref og.opcgroup);
                        varGroups.Add(og);
                    }
                    sdr.Close();
                }
                log.Success();
            }
            catch (Exception ex)
            {
                log.Fatal(ex);
            }
            return varGroups;
        }
        /* GROUPS AND ITEMS for variables in all WHERES!!!!!!!!!!!
         * Create the GROUPs for the OPC client   
         * base just on the WHERES clauses
         * Store groups in WhereGroups in array
         */
        private List<OPCGroup_t> InitWhereGroup()
        {
            string sqlQuery;
            sqlQuery = "select OpcName from PLC2DB_Tags, PLC2DB_TagsInWhere where PLC2DB_Tags.Tag = PLC2DB_TagsInWhere.Tag and len(rtrim(OpcName))>1 group by OpcName";
            Logging log = new Logging();
            List<OPCGroup_t> whereGroups = new List<OPCGroup_t>();
            try
            {
                using (SqlConnection connection = new SqlConnection(IntData.CfgConn))
                {
                    connection.Open();
                    SqlCommand cmd = new SqlCommand(sqlQuery, connection);
                    SqlDataReader sdr = cmd.ExecuteReader();
                    while (sdr.Read())
                    {
                        OPCGroup_t og = new OPCGroup_t
                        {
                            tagEvent = "",
                            opcgroup = new OPCController
                            {
                                ServerName = sdr[0].ToString(),
                                GroupName = "whereGrp" + whereGroups.Count
                            }
                        };
                        whereGroups.Add(og);
                        sqlQuery = "select PLC2DB_Tags.Tag as Tag, Description, PLCName, Register from PLC2DB_Tags, PLC2DB_TagsInWhere where OPCName  ='" + sdr[0].ToString() + "' and PLC2DB_Tags.Tag = PLC2DB_TagsInWhere.tag group by PLC2DB_Tags.Tag, Description, PLCName, Register";
                        CreateItemsFromSelect(sqlQuery, ref og.opcgroup);
                    }
                    sdr.Close();
                }
                log.Success();
            }
            catch (Exception ex)
            {
                log.Fatal(ex);
            }
            return whereGroups;
        }
        private int GetPeriod(int period, string unit)
        {
            switch (unit)
            {
                case "day": return period * 86400;
                case "hour": return period * 3600;
                case "min": return period * 60;
                case "sec": return period * 1;
                default: return period * 60;    //By default, take MINUTES!
            }
        }
        private List<List<TablesCommand>> MakeCommandsList(List<List<TablesCommand>> TablesCmd, string tagevent)
        {
            int tableNum = 0;
            int TotalCmds = 0;
            string sqlQuery;
            string[] flancStr = { "POSITIVE", "NEGATIVE" };
            string[] actionStr = { "INSERT", "UPDATE", "SELECT" };
            if (TablesCmd == null) TablesCmd = new List<List<TablesCommand>>();
            Logging log = new Logging();
            try
            {
                using (SqlConnection connection = new SqlConnection(IntData.CfgConn))
                {
                    connection.Open();
                    SqlCommand cmd = new SqlCommand()
                    {
                        Connection = connection
                    };
                    SqlDataReader sdr;

                    foreach (string flanc in flancStr)
                    {
                        List<TablesCommand> tableCommands = new List<TablesCommand>();
                        foreach (string action in actionStr)
                        {
                            sqlQuery = "select TableDest, cond from PLC2DB_Cfg where TagEvent = '" + tagevent + "' and " + flanc + " = '" + action + "' group by TableDest, cond";
                            log.Command = sqlQuery;
                            cmd.CommandText = sqlQuery;
                            sdr = cmd.ExecuteReader();
                            while (sdr.Read())
                            {
                                TablesCommand tableCmd = new TablesCommand();
                                if (++tableNum > TotalCmds)
                                {
                                    string whereaux = sdr[1].ToString().Trim();
                                    tableCmd.ColumnName = sdr[0].ToString().Trim();
                                    tableCmd.Command = action;
                                    tableCmd.Where = whereaux.Substring(0, 5) == "where" ? whereaux : "where " + whereaux;
                                    tableCmd.Columns = CreateColumnsList(tableCmd.ColumnName, tagevent, action, flanc, whereaux);
                                }
                                tableCommands.Add(tableCmd);
                            }
                            sdr.Close();
                            TotalCmds = TotalCmds > tableNum ? TotalCmds : tableNum;
                            log.Success();
                        }
                        TablesCmd.Add(tableCommands);
                    }
                }
                return TablesCmd;
            }
            catch (Exception ex)
            {
                log.Fatal(ex);
                return TablesCmd;
            }
        }
        private List<List<string>> FilluserProcedures(List<List<string>> ScriptPointer, string tagevent)
        {
            Logging log = new Logging();
            int TotalProc = 0;
            string[] flancStr = { "RunProcPos", "RunProcNeg" };
            try
            {
                foreach (string flanc in flancStr)
                {
                    int tableNum = 0;
                    string sqlQuery = "select " + flanc + " , cond from PLC2DB_Cfg where TagEvent = '" + tagevent + "' group by  " + flanc + " , cond ";
                    log.Command = sqlQuery;
                    using (SqlConnection connection = new SqlConnection(IntData.CfgConn))
                    {
                        connection.Open();
                        SqlCommand cmd = new SqlCommand(sqlQuery, connection);
                        SqlDataReader sdr = cmd.ExecuteReader();
                        List<string> Procedure = new List<string>();
                        while (sdr.Read())
                        {
                            if (sdr[0].ToString() != "" && ++tableNum > TotalProc)
                            {
                                Procedure.Add(sdr[0].ToString().Trim());
                            }
                        }
                        if (Procedure.Count != 0) ScriptPointer.Add(Procedure);
                        sdr.Close();
                        TotalProc = TotalProc > tableNum ? TotalProc : tableNum;
                    }
                    log.Success();
                }
                return ScriptPointer;
            }
            catch (Exception ex)
            {
                log.Fatal(ex);
                return ScriptPointer;
            }
        }
        private List<ColumnsCommand> CreateColumnsList(string TableName, string tagevent, string SQLInstruction, string flancStr, string Condition)
        {
            List<ColumnsCommand> Columns = new List<ColumnsCommand>();
            string KeepInTagStr = flancStr == "POSITIVE" ? "KeepInTagPos" : "KeepInTagNeg";
            string sqlQuery = "select Expression, ColumnDest, PLC2DB_Cfg.Enabled, Format, " + KeepInTagStr + ", WriteOnPLCTag,OPCName, PLCName, Register from PLC2DB_Cfg left outer join plc2db_tags on WriteOnPLCTag = Tag where TagEvent = '" + tagevent + "' and " + flancStr + " = '" + SQLInstruction + "' and TableDest ='" + TableName + "' and Cond ='" + Condition.Replace("'", "''") + "' ";
            Logging log = new Logging(sqlQuery);
            try
            {
                using (SqlConnection connection = new SqlConnection(IntData.CfgConn))
                {
                    connection.Open();
                    SqlCommand cmd = new SqlCommand(sqlQuery, connection);
                    SqlDataReader sdr = cmd.ExecuteReader();
                    while (sdr.Read())
                    {
                        ColumnsCommand column = new ColumnsCommand
                        {
                            Formula = sdr[0].ToString().Trim(),
                            name = sdr[1].ToString().Trim(),
                            Enabled = Convert.ToBoolean(sdr[2]),
                            Format = sdr[3].ToString().Trim(),
                            KeepInTag = sdr[4].ToString().Trim(),
                            WriteToTag = sdr[5].ToString().Trim(),
                            ServerName = sdr[6].ToString().Trim(),
                            PLCName = sdr[7].ToString().Trim(),
                            TagAddress = sdr[8].ToString().Trim()
                        };
                        Columns.Add(column);
                    }
                    sdr.Close();
                }
                log.Success();
            }
            catch (Exception ex)
            {
                log.Fatal(ex);
            }
            return Columns;
        }
        /*It creates all the items, for each group of where (group by OPCName)*/
        private bool CreateItemsFromSelect(string sqlquery, ref OPCController opcgroup)
        {
            Logging log = new Logging(sqlquery);
            try
            {
                using (SqlConnection connection = new SqlConnection(IntData.CfgConn))
                {
                    connection.Open();
                    SqlCommand cmd = new SqlCommand(sqlquery, connection);
                    SqlDataReader sdr = cmd.ExecuteReader();
                    while (sdr.Read())
                    {
                        string Tag = "";
                        string Register = "";
                        string Description = "";
                        string PLCName = "";
                        for (int i = 0; i < sdr.FieldCount; i++)
                        {
                            switch (sdr.GetName(i))
                            {
                                case "Tag": Tag = sdr[i].ToString().Trim(); break;
                                case "Register": Register = sdr[i].ToString().Trim(); break;
                                case "Description": Description = sdr[i].ToString().Trim(); break;
                                case "PLCName": PLCName = sdr[i].ToString().Trim(); break;
                            }
                        }
                        opcgroup.AddItem(Tag, Register, Description, PLCName);
                    }
                    if (opcgroup.ServerName.Length > 1) { opcgroup.Create(); }
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
        private void RefreshEventsGroup(TagHandler Tags)
        {
            for (int i = 0; i < EvGroups.Count; i++)
            {
                //All the events are in evGroups
                //refresh from OPC all events (diferents OPCs)
                if (EvGroups[i].opcgroup.ServerName.Length > 1)
                {
                    EvGroups[i].opcgroup.GetData();
                    Tags.RefreshFromGroup(EvGroups[i].opcgroup);
                }
                else
                {
                    Tags.RefreshFromEvents(EvGroups[i].opcgroup);
                }
            }
            //Refresh the struc EVENTS (that contain inserts/updates)
            for (int i = 0; i < OPCEvents.Count; i++)
            {
                int indexer = Tags.SearchTag(OPCEvents[i].Tagevent);
                if (indexer != -1)
                {
                    TagEvent evt = OPCEvents[i];
                    evt.Activated = Tags.Values[indexer].Activated;
                    evt.Deactivated = Tags.Values[indexer].Deactivated;
                    OPCEvents[i] = evt;
                }
            }
        }
        private void RefreshTimerEvents()
        {
            for (int i = 0; i < TimerEvents.Count; i++)
            {
                TimerEvent tte = TimerEvents[i];
                tte.Activated = false;
                if ((DateTime.Now - tte.LastTimeActivated).TotalSeconds >= (double)tte.TimerPeriod)
                {
                    tte.Activated = true;
                    tte.LastTimeActivated = DateTime.Now;
                }
                TimerEvents[i] = tte;
            }
        }
        private void RefreshVarGroup(string tagevent, TagHandler Tags)
        {
            for (int i = 0; i < VarGroups.Count; i++)
            {
                if (VarGroups[i].tagEvent == tagevent)
                {
                    VarGroups[i].opcgroup.GetData();
                    Tags.RefreshFromGroup(VarGroups[i].opcgroup);
                }
            }
        }
        private void RefreshWhereGroup(TagHandler Tags)
        {
            for (int i = 0; i < WhereGroups.Count; i++)
            {
                WhereGroups[i].opcgroup.GetData();
                Tags.RefreshFromGroup(WhereGroups[i].opcgroup);
            }
        }
    }
}
