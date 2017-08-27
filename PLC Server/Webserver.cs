using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Owin.Hosting;
using System.Timers;
using System.Management;
using System.IO.Ports;
using System.Configuration;

namespace AspNetSelfHostDemo
{
    public class WebServer
    {
        
        private IDisposable _webapp;
        private  Length_Reader lnthreader;
        public void Start()
        {
            try
            {

            #region serverstart
            _webapp = WebApp.Start<Startup>("http://localhost:8090");
                #endregion serverstart
                
                switch (PLCController.Initiate())
                {
                    case Serilportstatus.IsDetected:
                        break;
                    case Serilportstatus.IsError:
                        break;
                    case Serilportstatus.IsTryMax:
                        break;
                    default:
                        throw new Exception();
                        break;
                }


                #region get Activemachine 
                #endregion get Activemachine 
            }
            catch (Exception ex)
            {
                //wrote log forexception
               // throw;
            }
            finally
            {

            }

        }

        //private void timer_Elapsed(object sender, ElapsedEventArgs e)
        //{
        //    try
        //    {
        //        System.IO.StreamWriter wr = new System.IO.StreamWriter(@"G:\Akash\service test\out.txt", true);
        //        wr.WriteLine(DateTime.Now);
        //        //Console.WriteLine(DateTime.Now);
        //        wr.Close();
        //    }
        //    catch (Exception ex)
        //    {

        //        //throw;
        //    }

        //}

        public void Stop()
        {
            _webapp?.Dispose();
            //_timer?.Dispose();
        }
    }
}
