using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using PLCTools.Service;

namespace PLCTools.Common
{
    public class IntData
    {

        //public static bool readingOPC { get; set; } = false;
        public static bool IsOPCConnected { get; set; } = false;
        public static Boolean SaveRecordsetFlag { get; set; } = false;
        public static List<Table> TableList { get; set; } = new List<Table>();
        public static Queue<Logging> InforQ { get; set; } = new Queue<Logging>();
        public static Queue<Logging> WarningQ { get; set; } = new Queue<Logging>();
        public static Queue<Logging> ErrorQ { get; set; } = new Queue<Logging>();
        public static List<OPCController> OPCControllers { get; set; } = new List<OPCController>();

        public static string CfgConn { get; set; }
        public static string DestConn { get; set; }

        public const string QUEUE_NAME = "QUEUE_PLC2DB";
        public const string AUTO_INSERT_STR = "AutoIns";
        public const int timeInterval = 1000;
        private static List<Table> GetTablesList()
        {
            List<Table> TablesList = new List<Table>();
            Logging log = new Logging("select TableName, AutoInsert from PLC2DB_Tables");
            try
            {
                using (var connection = new SqlConnection(CfgConn))
                {
                    connection.Open();
                    SqlCommand cmd = new SqlCommand(log.Command, connection);
                    SqlDataReader sdr = cmd.ExecuteReader();
                    while (sdr.Read())
                    {
                        Table tb = new Table
                        {
                            Name = sdr[0].ToString(),
                            AutoInsert = Convert.ToInt32(sdr[1]) == 1
                        };
                        TablesList.Add(tb);
                    }
                    sdr.Close();
                }
                log.Success();

            }
            catch (Exception ex)
            {
                log.Fatal(ex);
            }
            return TablesList;
        }
        public static void InitializeData()
        {
            TableList = GetTablesList();
        }
    }
}
