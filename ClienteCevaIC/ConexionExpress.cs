using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;

namespace LogicaServicioCevaIC
{
    class ConexionExpress
    {
        public SqlConnection conectarExpress()
        {
            SqlConnection cn = null;
            cn = new SqlConnection(@"Persist Security Info=False;Integrated Security=SSPI; database=CEVA_ICTRACKING;server=DESKTOP-H3AF4NF;CURRENT LANGUAGE=SPANISH;");
            //cn = new SqlConnection(@"Persist Security Info=False;Integrated Security=SSPI; database=CEVA_ICTRACKING;server=WINCEVAPROD\SQLEXPRESS;CURRENT LANGUAGE=SPANISH;");
            return cn;
        }
    }
}
