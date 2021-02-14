using System;
using System.Collections.Generic;
using PLCTools.Common;
using OPCAutomation;
using System.ComponentModel;

namespace PLCTools.Models
{
    public class OPCController:IDisposable
    {
        private OPCServer opcgrp_server;
        private OPCGroup opcgrp_group;
        private Array opcgrp_arrayQualities = new int[0];
        private Array opcgrp_arrayHandles = new Int32[0];
        private Array opcgrp_arraySHandles = new int[0];
        private Array opcgrp_arrayErrors = new int[0];
        private Array opcgrp_arrayValues = new int[0];
        private Array opcgrp_arrayPaths = new string[0];
        private Array opcgrp_itemVar = new OPCItems[0];

        public int TotalItemNumber { get; set; } = 0;
        public string ServerName { get; set; }
        public string PLCName { get; set; }
        public string GroupName { get; set; }

        public OPCController()
        {
            Initialization();
        }
        public OPCController(string serverName, string plcName, string groupName)
        {
            ServerName = serverName;
            PLCName = plcName;
            GroupName = groupName;
            Initialization();
        }

        private void Initialization()
        {

        }
        public void Dispose()
        {
            opcgrp_server = null;
            opcgrp_arrayErrors = null;
            opcgrp_arrayHandles = null;
            opcgrp_arrayPaths = null;
            opcgrp_arrayQualities = null;
            opcgrp_arraySHandles = null;
            opcgrp_arrayValues = null;
            opcgrp_group = null;
            opcgrp_itemVar = null;
        }
        public void GetData(ref List<OPCItems> list)
        {
            foreach (OPCItems item in list)
            {
                AddItem(item.Tag, item.Address, item.Description, item.PLCName);
            }
            this.Create();
            this.GetData();
        }
        public void GetData(ref BindingList<OPCItems> list)
        {
            foreach (OPCItems item in list)
            {
                AddItem(item.Tag, item.Address, item.Description, item.PLCName);
            }
            this.Create();
            this.GetData();
        }

        public void PutData(List<OPCItems> list)
        {
            System.Array str = new object[list.Count + 1];
            foreach (OPCItems item in list)
            {
                AddItem(item.Tag, item.Address, item.Description, item.PLCName);
                str.SetValue(Convert.ToInt32(item.Value), list.IndexOf(item) + 1);
            }
            this.Create();
            this.PutData(str);
        }

        public Boolean AddItem(string newTag, string newAddress, string newDescprition, string newPLCName = null)
        {
            TotalItemNumber++;
            opcgrp_arrayPaths = ResizeArray(opcgrp_arrayPaths, TotalItemNumber + 1);
            opcgrp_arrayHandles = ResizeArray(opcgrp_arrayHandles, TotalItemNumber + 1);
            opcgrp_itemVar = ResizeArray(opcgrp_itemVar, TotalItemNumber + 1);
            opcgrp_arrayValues = ResizeArray(opcgrp_arrayValues, TotalItemNumber + 1);
            opcgrp_arrayErrors = ResizeArray(opcgrp_arrayErrors, TotalItemNumber + 1);
            opcgrp_arraySHandles = ResizeArray(opcgrp_arraySHandles, TotalItemNumber + 1);
            opcgrp_arrayQualities = ResizeArray(opcgrp_arrayQualities, TotalItemNumber + 1);
            OPCItems itemV;
            itemV = new OPCItems
            {
                Tag = newTag.Trim(),
                Address = newAddress.Trim(),
                Description = newDescprition.Trim(),
                PLCName = newPLCName == null ? PLCName.Trim() : newPLCName.Trim()
            };
            opcgrp_arrayHandles.SetValue(TotalItemNumber, TotalItemNumber);
            opcgrp_itemVar.SetValue(itemV, TotalItemNumber);
            opcgrp_arrayPaths.SetValue(newPLCName.Trim() + newAddress.Trim(), TotalItemNumber);
            return true;
        }
        public Boolean GetData(string tagName = "")
        {
            try
            {
                //for (; IntData.readingOPC; ) Thread.Sleep(10);
                if (IntData.IsOPCConnected)
                {
                    object timestamps = new object(); //store the timestamp of the read
                    object qualities = new object();
                    //IntData.readingOPC = true;
                    IntData.IsOPCConnected = true;
                    opcgrp_server = new OPCServer();
                    opcgrp_server.Connect(ServerName);
                    opcgrp_group = opcgrp_server.OPCGroups.Add(GroupName);
                    TotalItemNumber = opcgrp_arrayPaths.Length - 1;
                    opcgrp_group.OPCItems.DefaultIsActive = true;
                    opcgrp_group.OPCItems.AddItems(TotalItemNumber, opcgrp_arrayPaths, opcgrp_arrayHandles, out opcgrp_arraySHandles, out opcgrp_arrayErrors);
                    opcgrp_group.SyncRead((short)OPCAutomation.OPCDataSource.OPCDevice, TotalItemNumber, ref opcgrp_arraySHandles, out opcgrp_arrayValues, out opcgrp_arrayErrors, out qualities, out timestamps);
                    //IntData.readingOPC = false;
                    var qualitiesList = new List<int>();
                    var errorsList = new List<int>();
                    foreach (var quality in (Array)qualities) qualitiesList.Add(Convert.ToInt32(quality));
                    foreach (var error in opcgrp_arrayErrors) errorsList.Add(Convert.ToInt32(error));
                    for (int i = 1; i <= TotalItemNumber; i++)
                    {
                        OPCItems opcitem = new OPCItems();
                        opcitem = (OPCItems)opcgrp_itemVar.GetValue(i);
                        opcitem.Quality = qualitiesList[i - 1];
                        //opcitem.Error = errorsList[i - 1];
                        if (opcgrp_arrayValues.GetValue(i) is System.Int32)
                        {
                            opcitem.Value = (int)opcgrp_arrayValues.GetValue(i);
                        }
                        else if (opcgrp_arrayValues.GetValue(i) is System.Double)
                        {
                            opcitem.Value = (double)opcgrp_arrayValues.GetValue(i);
                        }
                        else
                        {
                            opcitem.Value = opcgrp_arrayValues.GetValue(i).ToString();
                        }
                        if (opcitem.Quality == 0)
                        {
                            IntData.IsOPCConnected = false;
                        }
                        opcgrp_itemVar.SetValue(opcitem, i);
                    }
                    if (!IntData.IsOPCConnected) throw new Exception("The PLC connection is broken!");
                }
                opcgrp_server.Disconnect();
                //log.Success();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }
        public Boolean PutData(Array Values)
        {
            IntData.IsOPCConnected = true;
            OPCServer OPCSvr = new OPCServer();
            try
            {
                IntData.IsOPCConnected = true;
                opcgrp_server = new OPCServer();
                opcgrp_server.Connect(ServerName);
                opcgrp_group = opcgrp_server.OPCGroups.Add(GroupName);
                TotalItemNumber = opcgrp_arrayPaths.Length - 1;
                opcgrp_group.OPCItems.DefaultIsActive = true;
                opcgrp_group.OPCItems.AddItems(TotalItemNumber, opcgrp_arrayPaths, opcgrp_arrayHandles, out opcgrp_arraySHandles, out opcgrp_arrayErrors);
                opcgrp_group.SyncWrite(TotalItemNumber, ref opcgrp_arraySHandles, ref Values, out opcgrp_arrayErrors);
                opcgrp_group.OPCItems.Remove(TotalItemNumber, ref opcgrp_arraySHandles, out opcgrp_arrayErrors);
                opcgrp_server.Disconnect();
                OPCSvr = null;
                return true;
            }
            catch (Exception ex)
            {
                IntData.IsOPCConnected = false;
                Console.WriteLine(ex.Message);
                return true;
            }
        }
        public Boolean Remove()
        {
            try
            {
                if (opcgrp_group != null)
                {
                    opcgrp_group = null;
                }
                if (opcgrp_server != null)
                {
                    opcgrp_server = null;
                }
                return true;
            }
            catch
            {
                return false;
            }
        }
        public Boolean Create()
        {
            try
            {
                IntData.IsOPCConnected = true;
                opcgrp_server = new OPCServer();
                opcgrp_server.Connect(ServerName);
                opcgrp_group = opcgrp_server.OPCGroups.Add(GroupName);
                TotalItemNumber = opcgrp_arrayPaths.Length - 1;
                opcgrp_group.OPCItems.AddItems(TotalItemNumber, opcgrp_arrayPaths, opcgrp_arrayHandles, out opcgrp_arraySHandles, out opcgrp_arrayErrors);
                opcgrp_server.Disconnect();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }
        public Boolean Activated(string tag)
        {
            int i = SearchTag(tag);
            if (i >= 0)
            {
                if (((OPCItems)opcgrp_itemVar.GetValue(i)).Activated())
                {
                    return true;
                }
            }
            return false;
        }
        public Boolean Deactivated(string tag)
        {
            int i = SearchTag(tag);
            if (i >= 0)
            {
                if (((OPCItems)opcgrp_itemVar.GetValue(i)).Deactivated())
                {
                    return true;
                }
            }
            return false;
        }
        public object GetTagValue(string tag)
        {
            int i = SearchTag(tag);
            if (i >= 0)
            {
                return Convert.ToInt32(((OPCItems)opcgrp_itemVar.GetValue(i)).Value);
            }
            else
            {
                return -1;
            }
        }
        public OPCItems GetTagItem(string tag)
        {
            int i = SearchTag(tag);
            if (i >= 0)
            {
                return (OPCItems)opcgrp_itemVar.GetValue(i);
            }
            else
            {
                return null;
            }
        }
        public string GetTagName(int i)
        {
            if (i >= 1 && i <= TotalItemNumber)
            {
                return ((OPCItems)opcgrp_itemVar.GetValue(i)).Tag.ToString().Trim();
            }
            else
            {
                return null;
            }
        }

        private int SearchTag(string tag)
        {
            for (int i = 1; i <= TotalItemNumber; i++)
            {
                if (((OPCItems)opcgrp_itemVar.GetValue(i)).Tag == tag)
                {
                    return i;
                }
            }
            return -1;
        }
        private Array ResizeArray(System.Array oldArray, int newSize)
        {
            int oldSize = oldArray.Length;
            System.Type elementType = oldArray.GetType().GetElementType();
            System.Array newArray = System.Array.CreateInstance(elementType, newSize);
            int preserveLength = System.Math.Min(oldSize, newSize);
            if (preserveLength > 0)
                System.Array.Copy(oldArray, newArray, preserveLength);
            return newArray;
        }
        private Boolean GetTagDefinition(string tag, ref string address, ref string description, ref string PLCName)
        {
            int i = SearchTag(tag);
            if (i >= 0)
            {
                address = ((OPCItems)opcgrp_itemVar.GetValue(i)).Address;
                description = ((OPCItems)opcgrp_itemVar.GetValue(i)).Description;
                PLCName = ((OPCItems)opcgrp_itemVar.GetValue(i)).PLCName;
                return true;
            }
            else
            {
                return false;
            }
        }

    }
}
