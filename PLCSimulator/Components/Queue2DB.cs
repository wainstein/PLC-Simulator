using System;
using System.Threading;
using System.Data.SqlClient;
using PLCTools.Service;
using PLCTools.Common;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace PLCTools.Components
{
    public class Queue2DB : Misc
    {
        public class QueryResult
        {
            public string Command { get; set; }
            public Boolean Completed { get; set; } = false;
            public int ExecuteTime { get; set; }
            public int AffectedRecordsets { get; set; }
            public Boolean IsAutoInsert { get; set; } = false;
            public IAsyncResult Result { get; set; }
            public int RetryCount { get; set; } = 0;
        }

        private int inTimer = 0;
        public Queue2DB(string _connTable)
        {
            try
            {
                if (_connTable != "") IntData.DestConn = _connTable;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + " Initialization failed");
            }
        }

        public void OnStart()
        {
            isConnecting = true;
            var t1 = Task.Run(() =>
            {
                while (isConnecting)
                {
                    if (Interlocked.Exchange(ref inTimer, 1) == 0)
                    {
                        MQHandler queue = new MQHandler();
                        string rcvd;
                        do
                        {
                            rcvd = queue.ReceiveMessage(IntData.QUEUE_NAME, IntData.QUEUE_NAME + " Message");//Dequeue the query
                            if (rcvd != "")
                            {
                                ThreadHandler.ThreadLocker();//Pause the program when thread is locked 
                                ExecuteSQL(rcvd);
                            }
                        } while (rcvd != "");
                        Interlocked.Exchange(ref inTimer, 0);
                    }
                    Thread.Sleep(1000);
                }
            });
        }
        private Boolean ExecuteSQL(string commandStr)
        {

            string tableName = "";
            string operation = "";
            bool autoInsert = false;
            SQLHandler.ParseTable(commandStr, ref tableName, ref operation);
            QueryResult query = new QueryResult();//Declare a struct to record the query.
            ThreadHandler.LockThread(); // Lock the thread here
            for (int attempts = 0; attempts < 5; attempts++)
            {
                try
                {
                    if (Log.Count > 200) Log.Dequeue();
                    using (SqlConnection connction = new SqlConnection(IntData.DestConn))
                    {
                        query.Command = commandStr;
                        connction.Open();
                        if (query.Command.Length > IntData.AUTO_INSERT_STR.Length)
                        {
                            if (query.Command.Substring(query.Command.Length - IntData.AUTO_INSERT_STR.Length, IntData.AUTO_INSERT_STR.Length) == IntData.AUTO_INSERT_STR)
                            {
                                query.Command = query.Command.Substring(0, query.Command.Length - IntData.AUTO_INSERT_STR.Length);
                                autoInsert = true;
                            }
                        }
                        SqlCommand command = new SqlCommand(query.Command, connction);
                        query.Result = command.BeginExecuteNonQuery();
                        query.ExecuteTime = SqlLocker(query.Result);
                        query.Completed = true;
                        query.AffectedRecordsets = command.EndExecuteNonQuery(query.Result);
                        Log.Enqueue(query.Command);//Queue the query.
                        if (autoInsert && query.AffectedRecordsets == 0)
                        {
                            command.CommandText = SQLHandler.GenerateInsert(query.Command);
                            query.Result = command.BeginExecuteNonQuery();
                            query.ExecuteTime = SqlLocker(query.Result);
                            query.IsAutoInsert = true;
                            query.Completed = true;
                            query.AffectedRecordsets = command.EndExecuteNonQuery(query.Result);
                            Log.Enqueue(query.Command);//Queue the qurey.
                        }
                    }
                    break;
                }
                catch (Exception ex)
                {
                    query.Completed = false;
                    query.Command += "\t Error:" + ex.Message;
                    Log.Enqueue(query.Command);
                    query.RetryCount++;
                }
            }
            //insert log here
            ThreadHandler.ReleaseThread(); // Release the thread here
            return true;
        }
        private int SqlLocker(IAsyncResult result, int timeInterval = 10)
        {
            int time = 0;
            while (!result.IsCompleted)
            {
                Thread.Sleep(timeInterval);
                time++;
            }
            return time * timeInterval;
        }
    }
}
