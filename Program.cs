using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace NotifivationCheckOutDate
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new Service1()
            };
            ServiceBase.Run(ServicesToRun);

            //debug
            //var service1 = new Service1();
            //service1.OnDebug();
            //System.Threading.Thread.Sleep(System.Threading.Timeout.Infinite);
        }
    }
}
