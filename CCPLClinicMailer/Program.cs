using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace CCPLClinicMailer
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            //===========Actual==============
            //            /*
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new ClinicMailer()
            };
            ServiceBase.Run(ServicesToRun);
            //*/

            //================ For Debugging======================
            //
            /*
            #if (!DEBUG)
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[] 
                { 
                        new PharmaMailer() 
                };
                ServiceBase.Run(ServicesToRun);
            #else
                ClinicMailer myServ = new ClinicMailer();
                myServ.Sale_Summary_Night();
            #endif

            //*/
        }
    }
}
