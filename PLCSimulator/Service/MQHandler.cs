using MSMQ;
using System;
using PLCTools.Common;
using PLCTools.Models;

namespace PLCTools.Service
{
    class MQHandler
    {        
        private Boolean CreatePrivateQueue(object varQueue)
        {
            Logging log = new Logging(varQueue.ToString());
            try
            {
                MSMQQueueInfo info = new MSMQQueueInfo
                {
                    PathName = ".\\PRIVATE$\\" + varQueue.ToString()
                };
                Boolean IsTransactional = false;
                Boolean IsWorldReadable = true;
                info.Create(IsTransactional, IsWorldReadable);
                log.Success();
                return true;
            }
            catch (Exception ex)
            {
                log.Fatal(ex);
                return false;
            }
        }
        private Boolean SendMessage(string varQueue, object varMessage)
        {
            Logging log = new Logging(varQueue+": "+ varMessage);
            try
            {
                MSMQQueueInfo info = new MSMQQueueInfo
                {
                    PathName = ".\\PRIVATE$\\" + varQueue.ToString()
                };
                MSMQQueue queue = info.Open((int)MQACCESS.MQ_SEND_ACCESS, (int)MQSHARE.MQ_DENY_NONE);
                MSMQMessage message = new MSMQMessage();
                message = new MSMQMessage
                {
                    Label = varMessage.ToString(),
                    Body = varMessage.ToString()
                };
                message.Send(queue);
                queue.Close();
                log.Success();
                return true;
            }
            catch (Exception ex)
            {
                log.Fatal(ex);
                return false;
            }
        }
        public string ReceiveMessage(object varQueue, object varMessageLabel)
        {
            Logging log = new Logging
            {
                Command = varQueue + ":" + varMessageLabel
            };
            try
            {
                string strMessageLabel = varMessageLabel.ToString();
                MSMQQueueInfo info = new MSMQQueueInfo
                {
                    PathName = ".\\PRIVATE$\\" + varQueue.ToString()
                };
                MSMQQueue queue = info.Open((int)MQACCESS.MQ_RECEIVE_ACCESS, (int)MQSHARE.MQ_DENY_NONE);
                MSMQMessage message = queue.Receive(ReceiveTimeout: 1000);                
                if (message != null)
                {
                    log.Success(Convert.ToString(message.Body));
                    return log.Information;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception ex)
            {
                log.Fatal(ex);
                return "";
            }
        }
        private Boolean DeletePrivateQueue(object varQueue)
        {
            MSMQQueueInfo info = new MSMQQueueInfo
            {
                PathName = ".\\PRIVATE$\\" + varQueue.ToString()
            };
            info.Delete();
            return true;
        }
        public Boolean CreateQueue()
        {
            CreatePrivateQueue(IntData.QUEUE_NAME);
            return true;
        }
        public void SendMsg(string cmdSend)
        {
            SendMessage(IntData.QUEUE_NAME, cmdSend);
        }
        public Boolean DeleteQueue()
        {
            DeletePrivateQueue(IntData.QUEUE_NAME);
            return true;
        }
    }
}
