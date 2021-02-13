using System;
using System.Net.Sockets;
using System.Text;
using System.Diagnostics;
using System.Net;

namespace PLCTools.Service
{
	/// <summary>
	/// Summary description for ClientMsg.
	/// </summary>
	public class ClientMsg : BaseClass
	{
		#region Declare instacne variables

		private static Encoding ASCII = Encoding.ASCII;

		// defaults
		private string server = "localhost";
		private int port = 9000;
		private bool loggedin = false;
		private Socket clientSocket = null;

		#endregion

		public ClientMsg()
		{
			//
			// TODO: Add constructor logic here
			//

			this.port = Int32.Parse(Panel.strSendPort);
			this.server = Panel.strServerIP;
		}

		public void SendMsgToServer(string data)
		{
			if (this.loggedin)
			{
				this.Close();
			}

			actionLog("ClientMsg:SendMsgToServer:", "Opening connection to " + this.server);

			System.Net.IPAddress addr = null;
			System.Net.IPEndPoint ep = null;

			try
			{
				this.clientSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);

				addr = Dns.Resolve(this.server).AddressList[0];
				ep = new System.Net.IPEndPoint(addr, this.port);

				this.clientSocket.Connect(ep);

				byte[] byteData = Encoding.ASCII.GetBytes(data);

				if (byteData.Length > 0)
				{
					this.clientSocket.Send(byteData);
					actionLog("ClientMsg:SendMsgToServer:", "Sent time message to server: " + data);
				}
			}

			catch (Exception ex)
			{
				if (this.clientSocket != null && this.clientSocket.Connected)
				{
					this.clientSocket.Close();
				}

				StackTrace st = new StackTrace(new StackFrame(true));
				actionLog("ClientMsg:SendMsgToServer(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Could not connect to remote server " + ex.ToString());
			}
		}

		public void Close()
		{
			actionLog("ClientMsg:Close:", "Closing connection to " + this.server);

			if (this.clientSocket != null)
			{
				this.clientSocket = null;
			}

			this.loggedin = false;
		}
	}
}
