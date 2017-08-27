using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;

namespace AspNetSelfHostDemo
{
    public enum Dbbitsize
    {

        bitdbsize8 = 8,
        bitdbsize16 = 16,
        bitdbsize32 = 32,
        bitdbsize7 = 7

    }

    public enum Serilportstatus
    {
       IsDetected=0,
       IsTryMax=1,
       IsError=2,
    }
    public class PLCserialConfig
    {
        public string _plcPortName { get; set; }
        public PLCserialConfig()
        {
            #region serialportDetection
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'");
                // var portnames = SerialPort.GetPortNames();
                //string st = 'thsiissd is';
                var ports = searcher.Get().Cast<ManagementBaseObject>().ToList().Select(p => p["Caption"].ToString());
                 string _plcPort = ports.Where(sp => sp.ToLower().Contains((ConfigurationManager.AppSettings["serialport"]))).FirstOrDefault();// portnames.w (n => n + " - " + ports.FirstOrDefault(s => s.Contains(n))).ToList();
                if (_plcPort?.Length > 0 && _plcPort.ToLower().LastIndexOf("(com") > 0)
                {
                    _plcPortName= _plcPort.Substring(_plcPort.ToLower().LastIndexOf("(com")+1, 4);
                }
                else
                {
                    _plcPortName = null;
                }
                //portList.FirstOrDefault().Substring(portList.FirstOrDefault().ToLower().LastIndexOf("com7"), 4)
                //  var portList = portnames.Select(n => n + " - " + ports.FirstOrDefault(s => s.Contains(n))).ToList();
            }
            catch (Exception ex)
            {

                throw;
            }
            #endregion serialportDetection
        }
    }
}
