using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO.Ports;

using System.Management;
using System.Management.Instrumentation;
using System.Timers;

namespace AspNetSelfHostDemo
{
    public static class PLCController
    {
        private static SerialPort sp = new SerialPort();
        private static Timer _timer = new Timer();
        public static List<Machine> lstMachine = new List<Machine>();
        public static string modbusStatus;
        static ManagementEventWatcher w = null;
        public static bool IsOpen { get; set; }
        public static string _portname { get; set; }

        #region Constructor / Deconstructor
        public static Serilportstatus Initiate()
        {
            //need to get list of port 
            try
            {

                int trycount = 0;
                AddRemoveUSBHandler();
                while (string.IsNullOrEmpty(_portname) || !sp.IsOpen)///run until get port
                {
                    _portname = new PLCserialConfig()._plcPortName;
                    if (!string.IsNullOrEmpty(_portname) && !sp.IsOpen)
                    {
                        IsOpen = Open(_portname, 19200, Dbbitsize.bitdbsize8.GetHashCode(), Parity.Even, StopBits.One);
                        //return Serilportstatus.IsDetected;

                        if (IsOpen)
                        {
                            lstMachine.Add(new Machine { machineID = 1, lengthRegiter = 1, MachineName = "test" });
                            _timer.Elapsed += new ElapsedEventHandler(Read_Length);
                            _timer.Interval = Convert.ToInt32(ConfigurationManager.AppSettings["lengthtimer"] ?? "2000");
                            _timer.Enabled = true;
                        }
                    }
                    else
                    {
                        IsOpen = false;
                    }
                    trycount++;
                    if (trycount > Convert.ToInt32(ConfigurationManager.AppSettings["DetectCounter"]))
                    {
                        return Serilportstatus.IsTryMax;
                    }
                }
            }
            catch (Exception ex)
            {
                return Serilportstatus.IsError;
                //throw;
            }
            return Serilportstatus.IsDetected;


        }
        #endregion

        #region Open / Close Procedures
        public static bool Open(string portName, int baudRate, int databits, Parity parity, StopBits stopBits)
        {

            //GetHashCode for list of port and opwn it 
            //Ensure port isn't already opened:
            if (!sp.IsOpen)
            {
                //Assign desired settings to the serial port:
                sp.PortName = portName;
                sp.BaudRate = baudRate;
                sp.DataBits = databits;
                sp.Parity = parity;
                sp.StopBits = stopBits;
                //These timeouts are default and cannot be editted through the class at this point:
                sp.ReadTimeout = 1000;
                sp.WriteTimeout = 1000;
                //sp.Handshake = Handshake.RequestToSend;

                try
                {
                    sp.Open();
                }
                catch (Exception err)
                {
                    modbusStatus = "Error opening " + portName + ": " + err.Message;
                    return false;
                }
                //modbusStatus = portName + " opened successfully";
                return true;
            }
            else
            {
                //modbusStatus = portName + " already opened";
                return false;
            }
        }
        public static bool Close()
        {
            //Ensure port is opened before attempting to close:
            if (sp.IsOpen)
            {
                try
                {
                    sp.Close();
                }
                catch (Exception err)
                {
                    modbusStatus = "Error closing " + sp.PortName + ": " + err.Message;
                    return false;
                }
                modbusStatus = sp.PortName + " closed successfully";
                return true;
            }
            else
            {
                modbusStatus = sp.PortName + " is not open";
                return false;
            }
        }
        #endregion

        #region CRC Computation
        private static void GetCRC(byte[] message, ref byte[] CRC)
        {
            //Function expects a modbus message of any length as well as a 2 byte CRC array in which to 
            //return the CRC values:

            ushort CRCFull = 0xFFFF;
            byte CRCHigh = 0xFF, CRCLow = 0xFF;
            char CRCLSB;

            for (int i = 0; i < (message.Length) - 2; i++)
            {
                CRCFull = (ushort)(CRCFull ^ message[i]);

                for (int j = 0; j < 8; j++)
                {
                    CRCLSB = (char)(CRCFull & 0x0001);
                    CRCFull = (ushort)((CRCFull >> 1) & 0x7FFF);

                    if (CRCLSB == 1)
                        CRCFull = (ushort)(CRCFull ^ 0xA001);
                }
            }
            CRC[1] = CRCHigh = (byte)((CRCFull >> 8) & 0xFF);
            CRC[0] = CRCLow = (byte)(CRCFull & 0xFF);
        }
        #endregion

        #region Build Message
        private static void BuildMessage(byte address, byte type, ushort start, ushort registers, ref byte[] message)
        {
            //Array to receive CRC bytes:
            byte[] CRC = new byte[2];

            message[0] = address;
            message[1] = type;
            message[2] = (byte)(start >> 8);
            message[3] = (byte)start;
            message[4] = (byte)(registers >> 8);
            message[5] = (byte)registers;

            GetCRC(message, ref CRC);
            message[message.Length - 2] = CRC[0];
            message[message.Length - 1] = CRC[1];
        }
        #endregion

        #region Check Response
        private static bool CheckResponse(byte[] response)
        {
            //Perform a basic CRC check:
            byte[] CRC = new byte[2];
            GetCRC(response, ref CRC);
            if (CRC[0] == response[response.Length - 2] && CRC[1] == response[response.Length - 1])
                return true;
            else
                return false;
        }
        #endregion

        #region Get Response
        private static void GetResponse(ref byte[] response)
        {
            //There is a bug in .Net 2.0 DataReceived Event that prevents people from using this
            //event as an interrupt to handle data (it doesn't fire all of the time).  Therefore
            //we have to use the ReadByte command for a fixed length as it's been shown to be reliable.
            lock (sp)
            {
                for (int i = 0; i < response.Length; i++)
                {
                    response[i] = (byte)(sp.ReadByte());
                }
            }
        }
        #endregion

        #region Function 16 - Write Multiple Registers
        public static bool SendFc16(byte address, ushort start, ushort registers, short[] values)
        {
            //Ensure port is open:
            if (sp.IsOpen)
            {
                //Clear in/out buffers:
                sp.DiscardOutBuffer();
                sp.DiscardInBuffer();
                //Message is0 1 addr + 1 fcn + 2 start + 2 reg + 1 count + 2 * reg vals + 2 CRC
                byte[] message = new byte[9 + 2 * registers];
                //Function 16 response is fixed at 8 bytes
                byte[] response = new byte[8];

                //Add bytecount to message:
                message[6] = (byte)(registers * 2);
                //Put write values into message prior to sending:
                for (int i = 0; i < registers; i++)
                {
                    message[7 + 2 * i] = (byte)(values[i] >> 8);
                    message[8 + 2 * i] = (byte)(values[i]);
                }
                //Build outgoing message:
                BuildMessage(address, (byte)16, start, registers, ref message);

                //Send Modbus message to Serial Port:
                try
                {
                    sp.Write(message, 0, message.Length);
                    GetResponse(ref response);
                }
                catch (Exception err)
                {
                    modbusStatus = "Error in write event: " + err.Message;
                    return false;
                }
                //Evaluate message:
                if (CheckResponse(response))
                {
                    modbusStatus = "Write successful";
                    return true;
                }
                else
                {
                    modbusStatus = "CRC error";
                    return false;
                }
            }
            else
            {
                modbusStatus = "Serial port not open";
                return false;
            }
        }
        #endregion

        #region Function 3 - Read Registers
        public static bool SendFc3(byte address, ushort start, ushort registers, ref short[] values)
        {
            //Ensure port is open:
            if (!IsOpen)
            {
                //Open();
            }
            if (sp.IsOpen)
            {
                //Clear in/out buffers:
                sp.DiscardOutBuffer();
                sp.DiscardInBuffer();
                //Function 3 request is always 8 bytes:
                byte[] message = new byte[8];
                //Function 3 response buffer:
                byte[] response = new byte[5 + 2 * registers];
                //Build outgoing modbus message:
                BuildMessage(address, (byte)3, start, registers, ref message);
                //Send modbus message to Serial Port:
                try
                {
                    lock(sp)
                    {
                        sp.Write(message, 0, message.Length);
                    }
                  
                    GetResponse(ref response);
                }
                catch (Exception err)
                {
                    modbusStatus = "Error in read event: " + err.Message;
                    return false;
                }
                //Evaluate message:
                if (CheckResponse(response))
                {
                    //Return requested register values:
                    for (int i = 0; i < (response.Length - 5) / 2; i++)
                    {
                        values[i] = response[2 * i + 3];
                        values[i] <<= 8;
                        values[i] += response[2 * i + 4];
                    }
                    //modbusStatus = "Read successful";
                    return true;
                }
                else
                {
                    //modbusStatus = "CRC error";
                    return false;
                }
            }
            else
            {
                modbusStatus = "Serial port not open";
                return false;
            }

        }
        #endregion

        #region Length_Reader
        private static void Read_Length(object sender, ElapsedEventArgs e)
        {
            try
            {

                if (IsOpen)
                {
                    foreach (Machine item in lstMachine)
                    {
                        try
                        {
                            PollFunction(item);
                        }
                        catch (Exception ex)
                        {
                            //machine level error
                            //throw;
                        }
                    }
                }
                else
                {
                    //blank entry in  database
                }

            }
            catch (Exception)
            {

                throw;
            }
        }

        #region Poll Function
        public static string  PollFunction(Machine machn)
        {


            //Create array to accept read values:

            ushort pollStart = 0;
            ushort pollLength = 1;

            short[] values = new short[pollLength];


            //Read registers and display data in desired format:
            int counter = 0;
            while (true)
            {
                
                if(counter>4)
                {

                    break;
                }
                try
                {
                    if(SendFc3(Convert.ToByte(machn.machineID.ToString()), pollStart, pollLength, ref values))
                    {
                        break;
                    }
                }
                catch (Exception err)
                {
                    //throw err;
                    //DoGUIStatus("Error in modbus read: " + err.Message);
                }
                counter++;
            }
            if(counter==5)
            {
                return "";
            }
            else
            {
                return values[0].ToString();
            }

           

            //values has response 

            //lstResponse.Add(machn.machineID, values[0].ToString());


            #region datatype
            //if (RegisterType.SelectedIndex == 1)
            //{
            //    switch (dataType)
            //    {

            //        case "Decimal":
            //            for (int i = 0; i < pollLength; i++)
            //            {
            //                itemString = "[" + Convert.ToString(pollStart + i + 40001) + "] , MB[" +
            //                    Convert.ToString(pollStart + i) + "] = " + values[i].ToString();
            //                DoGUIUpdate(itemString);
            //            }
            //            break;
            //        case "Hexadecimal":
            //            for (int i = 0; i < pollLength; i++)
            //            {
            //                itemString = "[" + Convert.ToString(pollStart + i + 40001) + "] , MB[" +
            //                    Convert.ToString(pollStart + i) + "] = " + values[i].ToString("X");
            //                DoGUIUpdate(itemString);
            //            }
            //            break;
            //        case "Float":
            //            for (int i = 0; i < (pollLength / 2); i++)
            //            {
            //                int intValue = (int)values[2 * i];
            //                intValue <<= 16;
            //                intValue += (int)values[2 * i + 1];
            //                itemString = "[" + Convert.ToString(pollStart + 2 * i + 40001) + "] , MB[" +
            //                    Convert.ToString(pollStart + 2 * i) + "] = " +
            //                    (BitConverter.ToSingle(BitConverter.GetBytes(intValue), 0)).ToString();
            //                DoGUIUpdate(itemString);
            //            }
            //            break;
            //        case "Reverse":
            //            for (int i = 0; i < (pollLength / 2); i++)
            //            {
            //                int intValue = (int)values[2 * i + 1];
            //                intValue <<= 16;
            //                intValue += (int)values[2 * i];
            //                itemString = "[" + Convert.ToString(pollStart + 2 * i + 40001) + "] , MB[" +
            //                    Convert.ToString(pollStart + 2 * i) + "] = " +
            //                    (BitConverter.ToSingle(BitConverter.GetBytes(intValue), 0)).ToString();
            //                DoGUIUpdate(itemString);
            //            }
            //            break;
            //    }
            //}
            //else
            //{

            //    switch (dataType)
            //    {

            //        case "Decimal":
            //            for (int i = 0; i < pollLength; i++)
            //            {
            //                itemString = "[" + Convert.ToString(pollStart + i) + "] , MB[" +
            //                    Convert.ToString(pollStart + i) + "] = " + values[i].ToString();
            //                DoGUIUpdate(itemString);
            //            }
            //            break;
            //        case "Hexadecimal":
            //            for (int i = 0; i < pollLength; i++)
            //            {
            //                itemString = "[" + Convert.ToString(pollStart + i) + "] , MB[" +
            //                    Convert.ToString(pollStart + i) + "] = " + values[i].ToString("X");
            //                DoGUIUpdate(itemString);
            //            }
            //            break;
            //        case "Float":
            //            for (int i = 0; i < (pollLength / 2); i++)
            //            {
            //                int intValue = (int)values[2 * i];
            //                intValue <<= 16;
            //                intValue += (int)values[2 * i + 1];
            //                itemString = "[" + Convert.ToString(pollStart + 2 * i) + "] , MB[" +
            //                    Convert.ToString(pollStart + 2 * i) + "] = " +
            //                    (BitConverter.ToSingle(BitConverter.GetBytes(intValue), 0)).ToString();
            //                DoGUIUpdate(itemString);
            //            }
            //            break;
            //        case "Reverse":
            //            for (int i = 0; i < (pollLength / 2); i++)
            //            {
            //                int intValue = (int)values[2 * i + 1];
            //                intValue <<= 16;
            //                intValue += (int)values[2 * i];
            //                itemString = "[" + Convert.ToString(pollStart + 2 * i) + "] , MB[" +
            //                    Convert.ToString(pollStart + 2 * i) + "] = " +
            //                    (BitConverter.ToSingle(BitConverter.GetBytes(intValue), 0)).ToString();
            //                DoGUIUpdate(itemString);
            //            }
            //            break;
            //    }
            //}
            #endregion datatype
        }
        #endregion

        #endregion Length_Reader

        #region usb
       
        public static void AddRemoveUSBHandler()
        {
            WqlEventQuery IN;
            WqlEventQuery OUT;
            ManagementScope scope = new ManagementScope("root\\CIMV2");
            scope.Options.EnablePrivileges = true;
            try
            {
                OUT = new WqlEventQuery();
                OUT.EventClassName = "__InstanceDeletionEvent";
                OUT.WithinInterval = new TimeSpan(0, 0, 3);
                OUT.Condition = @"TargetInstance ISA 'Win32_USBHub'";
                w = new ManagementEventWatcher(scope, OUT);
                w.EventArrived += new EventArrivedEventHandler(USBRemoved);
                w.Start();


               IN = new WqlEventQuery();
               IN.EventClassName = "__InstanceCreationEvent";
               IN.WithinInterval = new TimeSpan(0, 0, 3);
                IN.Condition = @"TargetInstance ISA 'Win32_USBHub'";
                w = new ManagementEventWatcher(scope, IN);
                w.EventArrived += new EventArrivedEventHandler(USBAdded);
                w.Start();
            }

            catch (Exception e)
            {

                Console.WriteLine(e.Message);
                if (w != null)
                    w.Stop();
            }
        }

   

        public static void USBAdded(object sender, EventArgs e)
        {
            
        }

        public static void USBRemoved(object sender, EventArgs e)
        {
           
        }
        #endregion usb

    }



}
