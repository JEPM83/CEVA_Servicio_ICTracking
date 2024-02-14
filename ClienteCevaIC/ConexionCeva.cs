using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace LogicaServicioCevaIC
{
     class ConexionCeva
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Inventory());
        }

        public SqlConnection conectarSap() {
            SqlConnection cn = null;
            cn = new SqlConnection(@"Persist Security Info=False;User ID=sa;Password=123;Initial Catalog=SBO_BDINT;Server=SUPERPC\SUPERPC");
            //cn = new SqlConnection(@"Persist Security Info=False;User ID=Consultor1;Password=Consult102%;Initial Catalog=SBO_BDINT;Server=192.168.1.201");
            return cn;
        }
    }

    
}
