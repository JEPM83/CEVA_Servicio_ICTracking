using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Timers;
using LogicaServicioCevaIC;

namespace ServicioWindows
{
    public partial class ServiceCeva : ServiceBase
    {
        LogicaServicioCevaIC.GetConfiguration logicanegocio = new LogicaServicioCevaIC.GetConfiguration();
        Timer tmServicio = null;
        private BackgroundWorker worker;
        public ServiceCeva()
        {
            InitializeComponent();
            worker = new BackgroundWorker();
            worker.DoWork += worker_Dowork;
            //tmServicio = new Timer(10000);
            tmServicio = new Timer(1000);

            tmServicio.Elapsed += new ElapsedEventHandler(tmServicio_Elapsed);
            ServiceName = "CevaICTrackingService";
        }

        void tmServicio_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (!worker.IsBusy)
                worker.RunWorkerAsync();
        }

        void worker_Dowork(object sender, DoWorkEventArgs e)
        {
            string hostgroupid = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            logicanegocio.ejecucion("LLL", "POSTADVANCEDSHIPMENTNOTICE", hostgroupid);

            hostgroupid = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            logicanegocio.ejecucion("LLL", "POSTSHIPMENTORDER", hostgroupid);

            hostgroupid = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            logicanegocio.ejecucion("GPA", "POSTADVANCEDSHIPMENTNOTICE", hostgroupid);

            hostgroupid = DateTime.Now.ToString("yyyyMMddHHmmssfff");
           logicanegocio.ejecucion("SCH", "POSTLOADXMLSHIPMENT", hostgroupid);
        }

        protected override void OnStart(string[] args)
        {
            //System.Diagnostics.Debugger.Launch();
            tmServicio.Start();
        }

        protected override void OnStop()
        {
            tmServicio.Stop();
        }
    }
}
