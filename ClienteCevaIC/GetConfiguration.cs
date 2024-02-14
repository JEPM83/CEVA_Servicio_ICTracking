using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Reflection;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using Renci.SshNet;
using Renci.SshNet.Common;
using Renci.SshNet.Sftp;
using System.Xml;
using Ionic.Zip;
using System.Globalization;
using System.Runtime.InteropServices;
using OfficeOpenXml;


namespace LogicaServicioCevaIC
{
    public class GetConfiguration
    {
        public string Server;
        public int Port;
        public string Route;
        public bool Sftp;
        public string Historic;
        public string UserSap;
        public string Password;
        public bool Email;
        public bool Attached;
        public int Code;
        public string Modelo;
        List<List<int>> KeySolum = new List<List<int>>();

        public void limpiar_variables()
        {
            //origen = null;
            Prefix = null;
            Extent = null;
            Separator = null;
            Server = null;
            Route = null;
            Historic = null;
            UserSap = null;
            Password = null;
            Modelo = null;
            Destino = null;
        }

        public void ejecucion(string cliente, string proceso,string hostgroupid) {
            limpiar_variables();
            if (cargar_configuracion(cliente, proceso))
            {
                cargar_ruta(cliente, proceso,hostgroupid);
            }
            else
            {
                //MessageBox.Show("Proceso no configurado o desactivado");
                //Inscribir error
            }
        }

        private bool cargar_configuracion(string cliente,string proceso)
        {
            SqlConnection cn = null;
            SqlCommand cmd = null;
            SqlDataReader rd = null;
            ConexionExpress conE = new ConexionExpress();
            int i = 0;
            try
            {
                cn = conE.conectarExpress();
                cn.Open();
                cmd = new SqlCommand("select Code from [@IC_PROCESO] where Code = '" + proceso + "' and state = 0", cn);
                rd = cmd.ExecuteReader();
                while (rd.Read())
                {
                    i++;
                }
                rd.Close();
                cn.Close();
                if (i == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            finally
            {
                rd.Close();
                cn.Close();
            }
        }

        private string cargar_cliente(string cliente)
        {
            SqlConnection cn = null;
            SqlCommand cmd = null;
            SqlDataReader rd = null;
            ConexionExpress conE = new ConexionExpress();
            string Ccliente = null;
            try
            {
                cn = conE.conectarExpress();
                cn.Open();
                cmd = new SqlCommand("select Code, Class, Name from [@ic_client] where state = 0 and code = '" + cliente + "'", cn);
                rd = cmd.ExecuteReader();
                while (rd.Read()) {
                    Ccliente = rd.GetValue(2).ToString(); 
                }
                rd.Close();
                cn.Close();
                return Ccliente;
            }
            finally
            {
                rd.Close();
                cn.Close();
            }
        }

        private void cargar_ruta(string cliente, string proceso,string hostgroupid)
        {
            SqlConnection cn = null;
            SqlCommand cmd = null;
            SqlDataReader rd = null;
            ConexionExpress conE = new ConexionExpress();
            int i = 0;
            string LogErr;
            bool Resultado = false;
            try
            {
                cn = conE.conectarExpress();
                cn.Open();
                cmd = new SqlCommand("select Code,Server,Port,Route,Sftp,Historic,UserSap,Password,Email,Attached,U_IC_MODELO,Zip,Subject,Environment,LogErr from [@IC_RUTA] where U_IC_CLIENTE = '" + cliente + "' and U_IC_PROCESO = '" + proceso + "' and state = 0 order by U_IC_PROCESO, Sequence", cn);
                rd = cmd.ExecuteReader();
                while (rd.Read())
                {
                    i++;
                    Code = int.Parse(rd.GetValue(0).ToString());
                    Server = rd.GetValue(1).ToString();
                    Port = string.IsNullOrEmpty(rd.GetValue(2).ToString()) ?  0 : int.Parse(rd.GetValue(2).ToString());
                    Route = rd.GetValue(3).ToString();
                    Sftp = rd.GetBoolean(4);
                    Historic = rd.GetValue(5).ToString();
                    UserSap = rd.GetValue(6).ToString();
                    Password = rd.GetValue(7).ToString();
                    Email = rd.GetBoolean(8);
                    Attached = rd.GetBoolean(9);
                    LogErr = rd.GetValue(14).ToString();

                    //Downloadfileftp();

                    Resultado = cargar_file(cliente, proceso, Code, Route, Historic, LogErr, hostgroupid, Email, Attached);

                    //if (Sftp == true)
                    //{
                    //    //Programación mas adelante
                    //}
                    //else {

                    //}

                    if (Email == true)
                    {
                        //if (Resultado == true)
                        //{
                            sendemail(cliente, Email, Code, Attached, "", hostgroupid, proceso, "");
                        //}
                    }
                }
                
                if (i == 0)
                {
                    //MessageBox.Show("RUTA no configurada o desactivada en proceso de INVENTORY");
                    //Inscribir error
                }
                rd.Close();
                cn.Close();
            }
            finally
            {
                rd.Close();
                cn.Close();
            }
        }

        //string origen;
        string Prefix;
        string Extent;
        string Separator;
        string Destino;
        private bool cargar_file(string cliente,string proceso, int PK,string Ruta,string Historico,string LogErr,string hostgroupid,bool email,bool attached)
        {
            SqlConnection cn = null;
            SqlCommand cmd = null;
            SqlDataReader rd = null;
            ConexionExpress conE = new ConexionExpress();
            
            string Ordenamiento;
            int Type = 0;
            bool Resultado = false;
            bool Resultado_Final = false;
            int i = 0;
            int Code_File = 0;
            int xHeader = 0;
            bool SDistinct = false;
            string SAttribute = null;
            try
            {
                cn = conE.conectarExpress();
                cn.Open();
                cmd = new SqlCommand("select Prefix,ICtent,Separator,Destino,[Order],Type,Code,isnull(Header,0),isnull(SDistinct,0),SAttribute from [@IC_FILE] where U_IC_RUTA = " + PK + " and state = 0 order by Code,[Order]", cn);
                rd = cmd.ExecuteReader();
                List<List<string>> lista = new List<List<string>>();
                while (rd.Read())
                {
                    i++;
                    Prefix = rd.GetValue(0).ToString();

                    //if (Prefix == "C1 - Order Header") {
                    //    var t = 1;   
                    //}

                    Extent = rd.GetValue(1).ToString();
                    Separator = rd.GetValue(2).ToString();
                    Destino = rd.GetValue(3).ToString();
                    Ordenamiento = rd.GetValue(4).ToString();
                    Type = int.Parse(rd.GetValue(5).ToString());
                    Code_File = int.Parse(rd.GetValue(6).ToString());
                    xHeader = int.Parse(rd.GetValue(7).ToString());
                    SDistinct = bool.Parse(rd.GetValue(8).ToString());
                    SAttribute = rd.GetValue(9).ToString();
                    if (String.IsNullOrEmpty(Ordenamiento))
                    {
                        Resultado = obtener_datos(cliente, proceso, Prefix, Extent, Separator, Ruta, Historico, Destino, LogErr, hostgroupid,0,"",email,attached,Code_File,Type,xHeader,SDistinct, SAttribute,PK);
                        if (Resultado == true) {
                            Resultado_Final = true;
                        }
                    }
                    else {
                        lista.Add(new List<string> { Prefix, Extent, Separator, Destino, Ordenamiento , Code_File.ToString(),SAttribute,Type.ToString()});
                    }
                }
                if (i == 0)
                {
                    //MessageBox.Show("INTERFAZ no configurada o desactivada en proceso de INVENTORY");
                    //Inscribir error
                }
                else if (i == 1 && lista.Count() > 0)
                {
                    Resultado = obtener_datos(cliente, proceso, Prefix, Extent, Separator, Ruta, Historico, Destino, LogErr, hostgroupid,0,"",email,attached,Code_File,Type,xHeader,SDistinct, SAttribute,PK);
                    if (Resultado == true)
                    {
                        Resultado_Final = true;
                    }
                }
                else if (i > 1 && lista.Count()>0) {
                    //string namefile;
                    List<string> detalle_lista = new List<string>();
                    detalle_lista = lista[0];
                    string[] files = Directory.GetFiles(Ruta, detalle_lista[0] + "*." + detalle_lista[1]);
                    
                    foreach (string names in files) {
                        Resultado = recorrer_grupo_archivo(cliente,proceso, Ruta, Historico,lista, Path.GetFileName(names),LogErr,hostgroupid,email,attached,Code_File,Type,xHeader,SDistinct, SAttribute,PK);
                        if (Resultado == true)
                        {
                            Resultado_Final = true;
                        }
                    }
                }
                rd.Close();
                cn.Close();
                return Resultado_Final;
            }
            finally
            {
                rd.Close();
                cn.Close();
            }
        }

        private bool recorrer_grupo_archivo(string cliente,string proceso,string ruta,string historico,List<List<string>> lista,string nombre,string logerr,string hostgroupid,bool email, bool attached,int code_file,int Type,int xHeader,bool SDistinct,string SAttribute,int pk) {
            List<List<string>> lista_trabajo = new List<List<string>>();
            string name_prefijo = null;
            string nombre_final = null;
            string[] files;
            int i = 0;
            bool estado = true;
            bool Resultado = false;
            bool Resultado_Final = false;
            foreach (List<string> detalle_lista in lista)
            {
                if (i == 0)
                {
                    files = Directory.GetFiles(ruta, nombre);
                }
                else {
                    nombre_final = detalle_lista[0] + name_prefijo;
                    files = Directory.GetFiles(ruta, nombre_final);
                    if (files.Count() == 0) {
                        estado = false;
                        break;
                    }
                }
                    name_prefijo = Path.GetFileName(files[0]).Substring(detalle_lista[0].Length, Path.GetFileName(files[0]).Length - detalle_lista[0].Length);
                    lista_trabajo.Add(new List<string> { detalle_lista[0], name_prefijo, detalle_lista[2], detalle_lista[3], detalle_lista[4] , detalle_lista[5] ,detalle_lista[6],detalle_lista[7]});
                i++;
            }
            if (estado == true) {
                foreach (List<string> procesar_lista in lista_trabajo) {
                    Resultado = obtener_datos(cliente, proceso, procesar_lista[0], procesar_lista[1], procesar_lista[2], ruta,historico , procesar_lista[3],logerr,hostgroupid, 1,procesar_lista[4],email,attached, int.Parse(procesar_lista[5]), int.Parse(procesar_lista[7]),xHeader,SDistinct, procesar_lista[6],pk);
                    if (Resultado == true)
                    {
                        Resultado_Final = true;
                    }
                }
                KeySolum.Clear();
            }
            return Resultado_Final;
        }

        public void sendemail(string cliente, bool archivo, int PK, bool adjunto,string Ssubject,string hostgroupid,string proceso,string origen)
        {
            if (archivo == true)
            {

                SqlConnection cn = null;
                SqlCommand cmd = null;
                SqlCommand cmd_log = null;
                SqlDataReader rd = null;
                SqlDataReader rd_log = null;
                ConexionExpress conE = new ConexionExpress();
                System.Data.DataTable dtSap;
                int i = 0;
                int j = 0;
                int z = 0;
                try
                {

                    cn = conE.conectarExpress();
                    cn.Open();
                    cmd = new SqlCommand("select Email,Password,Type from [@IC_CONTACT] where U_IC_RUTA = " + PK + " and state = 0", cn);
                    rd = cmd.ExecuteReader();
                    dtSap = new System.Data.DataTable() { TableName = "Contact" };

                    System.Net.Mail.MailMessage mmsg = new System.Net.Mail.MailMessage();
                    System.Net.Mail.SmtpClient clienteC = new System.Net.Mail.SmtpClient();
                    mmsg.Subject = "OPERADOR CEVA: Tablas intermedias - Proceso: (" + proceso + ") - " + Ssubject + " " + cargar_cliente(cliente) + " " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    mmsg.SubjectEncoding = System.Text.Encoding.UTF8;
                    while (rd.Read())
                    {
                        i++;

                        if (int.Parse(rd.GetValue(2).ToString()) == 0)
                        {
                            j++;
                            //Correo emisor
                            mmsg.From = new System.Net.Mail.MailAddress(rd.GetValue(0).ToString());
                            //Credenciales
                            clienteC.Credentials = new System.Net.NetworkCredential(rd.GetValue(0).ToString(), rd.GetValue(1).ToString());
                            //Si es gmail
                            clienteC.Port = 587;
                            clienteC.EnableSsl = true;
                            //
                            clienteC.Host = "smtp.gmail.com";
                        }
                        else if (int.Parse(rd.GetValue(2).ToString()) == 1)
                        {
                            mmsg.To.Add(rd.GetValue(0).ToString());
                        }
                        else if (int.Parse(rd.GetValue(2).ToString()) == 2)
                        {
                            mmsg.CC.Add(rd.GetValue(0).ToString());
                        }
                        else if (int.Parse(rd.GetValue(2).ToString()) == 3)
                        {
                            mmsg.Bcc.Add(rd.GetValue(0).ToString());
                        }
                    }
                    rd.Close();
                    cn.Close();
                    //Lista de interfaces

                    cn = conE.conectarExpress();
                    cn.Open();
                    //cmd_log = new SqlCommand("select isnull([LogErr] ,'') , isnull(FileName,''),MessageSystem  from [@IC_LOG] where HostGroupId = '" + hostgroupid +"' and State = 0", cn);
                    cmd_log = new SqlCommand("select isnull([LogErr] ,'') , isnull(FileName,''),isnull(MessageSystem,''),isnull([Historic],''),State  from [@IC_LOG] where HostGroupId = '" + hostgroupid + "'", cn);
                    rd_log = cmd_log.ExecuteReader();
                    string mensaje = null;
                    while (rd_log.Read()) {
                        //Environment.NewLine
                        if (adjunto == true) {
                            z++;
                            if (bool.Parse(rd_log.GetValue(4).ToString()) == false)
                            {
                                mensaje = mensaje + "ARCHIVO: " + rd_log.GetString(0) + rd_log.GetString(1) + "  MENSAJE: " + rd_log.GetString(2) + Environment.NewLine;
                                System.Net.Mail.Attachment attachment = new System.Net.Mail.Attachment(rd_log.GetString(0) + rd_log.GetString(1));
                                mmsg.Attachments.Add(attachment);
                            }
                            else {
                                mensaje = mensaje + "ARCHIVO: " + rd_log.GetString(3) + rd_log.GetString(1) + "  MENSAJE: " + rd_log.GetString(2) + Environment.NewLine;
                                System.Net.Mail.Attachment attachment = new System.Net.Mail.Attachment(rd_log.GetString(3) + rd_log.GetString(1));
                                mmsg.Attachments.Add(attachment);
                            }
                            
                        }
                    }
                    //
                    mmsg.Body = "Se ha procesado el siguiente proceso de importaciones: " + Environment.NewLine + Environment.NewLine + mensaje + Environment.NewLine + "Identificador: " + hostgroupid; 
                    
                    mmsg.BodyEncoding = System.Text.Encoding.UTF8;
                    mmsg.IsBodyHtml = false;
                    mmsg.Priority = System.Net.Mail.MailPriority.Normal;
                   
                    if (j == 1 && z > 0)
                    {
                        try
                        {
                            clienteC.Send(mmsg);
                        }
                        catch (System.Net.Mail.SmtpException ex)
                        {
                            RegistraEvento(hostgroupid, cliente, proceso, "", "", 0, ex.Message.ToString(), "Error en envio de correo: " + ex.Message.ToString(), PK);
                            //MessageBox.Show(ex.Message.ToString());
                            //Inscribir error
                        }
                        if (i == 0)
                        {
                            //MessageBox.Show("DESTINATARIOS no configurados o desactivados en proceso de INVENTORY");
                            //Inscribir error
                        }
                    }
                    else
                    {
                        //MessageBox.Show("Correo emisor no configurado en proceso INVENTORY");
                        //Inscribir error
                    }
                    ////
                    rd.Close();
                    cn.Close();
                    rd_log.Close();
                }
                catch(Exception e)
                {
                    rd.Close();
                    cn.Close();
                }
            }
        }

        private List<List<string>> ObtenerHeaders(int xHeader) {
            List<List<string>> listHeader = new List<List<string>>();
            return listHeader;
        }

        public class ArrayXml { 
            public List<DetalleArrayXml> Detalle { get; set; }
        }
        public class DetalleArrayXml { 
            public string nodo { get; set; }
            public List<xDetalleArrayXml> Dvalor { get; set; }
        }
        public class xDetalleArrayXml {
            public string nodo { get; set; }
            //public string valor { get; set; }
            public List<xSubDetalleArrayXml> DDvalor { get; set; }
        }
        public class xSubDetalleArrayXml {
            public string nodo { get; set; }
            public string valor { get; set; }
        }
        public ArrayXml GetXMLAsString(XmlDocument myxml)
        {
            XmlDocument xml = new XmlDocument();
            ArrayXml arreglo = new ArrayXml();
            arreglo.Detalle = new List<DetalleArrayXml>();
            DetalleArrayXml dl = new DetalleArrayXml();
            dl.Dvalor = new List<xDetalleArrayXml>();
            xDetalleArrayXml dll = new xDetalleArrayXml();
            dll.DDvalor = new List<xSubDetalleArrayXml>();
            xSubDetalleArrayXml dld = null;
            foreach (XmlNode node in myxml.DocumentElement.ChildNodes) {
                if (node.HasChildNodes) {
                    foreach(XmlNode x in node) {
                        xml.LoadXml("<detailroot>" + x.InnerXml + "</detailroot>");
                        dl = new DetalleArrayXml();
                        dl.nodo = x.Name.ToString();
                        dl.Dvalor = new List<xDetalleArrayXml>();
                        foreach (XmlNode nodeD in xml) {
                            foreach (XmlNode y in nodeD) {
                                dll = new xDetalleArrayXml();
                                dll.nodo = y.Name.ToString();
                                dll.DDvalor = new List<xSubDetalleArrayXml>();
                                foreach (XmlNode z in y) {
                                    dld = new xSubDetalleArrayXml();
                                    dld.nodo = z.Name.ToString();
                                    dld.valor = z.InnerText.ToString();
                                    dll.DDvalor.Add(dld);
                                }
                                dl.Dvalor.Add(dll);
                            }
                        }
                        arreglo.Detalle.Add(dl);
                    }
                    ////////CODIGO NO VALIDO
                    ////////for (int i = 0; i < node.ChildNodes.Count; i++)
                    ////////{
                    ////////    dl = new DetalleArrayXml();
                    ////////    dl.nodo = node.ChildNodes[i].Name.ToString();
                    ////////    dl.valor = node.ChildNodes[i].InnerXml.ToString();
                    ////////    //dl.valor = String.IsNullOrEmpty(node.ChildNodes[i].InnerXml) ? "" : node.ChildNodes[i].InnerXml.ToString() + "";
                    ////////    arreglo.Detalle.Add(dl);
                    ////////}
                }
            }
            return arreglo;
        }
        private bool obtener_datos(string cliente, string proceso,string prefijo, string ext, string separador,string ruta,string historico,string destino_file,string LogErr,string hostgroupid,int modo,string ordenamiento,bool email, bool attached,int code_file,int Type,int xHeader,bool SDistinct,string SAttribute,int pk)
        {
            //try
            //{
            //string mensaje = null;
                bool Resultado = false;

            //

            SqlConnection cn = null;
            SqlCommand cmd_datos = null;
            SqlCommand cmd_conf = null;
            SqlDataReader rd_conf = null;
            SqlDataReader rd_datos = null;
            ConexionExpress conE = new ConexionExpress();
            string sql_campos = null;
            string sql_values = null;
            string sql_where = null;
            string sql_distinct = null;
            string sql = null;

            if (SDistinct == true)
            {
                sql_distinct = " distinct ";
            }
            else {
                sql_distinct = null;
            }

            string coma = null;
            int j = 0;
            int z = 0;


            //

            if (Type == 0)
            {
                string[] files;
                if (modo == 0)
                {
                    files = Directory.GetFiles(ruta, prefijo + "*." + ext);
                }
                else
                {
                    files = Directory.GetFiles(ruta, prefijo + ext);
                }
                //if (files.Count() == 0)
                //{
                //    Resultado = true;
                //}
                //else {
                //    Resultado = false;
                //}
                foreach (string names in files)
                {

                    //int q = 0;
                    string namefile = Path.GetFileName((names));

                    try
                    {
                        if (Path.GetExtension(namefile).ToUpper() == ".TXT")
                        {

                            using (
                                System.IO.StreamReader file = new StreamReader(ruta + namefile,
                                    true))
                            {
                                string line = null;
                                sql = null;
                                int i = 0;
                                sql = "insert into " + destino_file;
                                sql_values = null;
                                while (!file.EndOfStream)
                                {
                                    if ((line = file.ReadLine()) != null)
                                    {
                                        string[] words = line.Split('\t');
                                        j = 0;
                                        foreach (string s in words)
                                        {
                                            if (string.IsNullOrEmpty(sql_values))
                                            {
                                                if (!string.IsNullOrEmpty(s.Trim()))
                                                {
                                                    sql_values = " select " + "'" + hostgroupid + "'," + "'" + s.ToString().Replace("'", "''").Trim() + "'";
                                                }
                                                else
                                                {
                                                    sql_values = " select " + "'" + hostgroupid + "'," + "null";
                                                }

                                            }
                                            else
                                            {
                                                if (!string.IsNullOrEmpty(s.Trim()))
                                                {
                                                    sql_values = " select " + "'" + hostgroupid + "'," + "'" + s.ToString().Replace("'", "''").Trim() + "'";
                                                }
                                                else
                                                {
                                                    sql_values = sql_values + ",'" + hostgroupid + "'," + "null";
                                                }

                                            }
                                            j = j + 1;
                                        }

                                        if (i == 0)
                                        {
                                            sql = sql + sql_values;
                                            sql_values = string.Empty;
                                        }
                                        if (i >= 1)
                                        {
                                            sql = sql + " union all " + sql_values;
                                            sql_values = string.Empty;
                                        }
                                        i = i + 1;
                                    }
                                }

                                EjecutarScript(sql);
                                //System.Console.WriteLine(sql);
                                file.Close();
                            }
                        }
                        else if (Path.GetExtension(namefile).ToUpper() == ".XLSX")
                        {

                            String sql_values_total = String.Empty;
                            FileInfo uploaded = new FileInfo(ruta + namefile);
                            using (ExcelPackage excel = new ExcelPackage(uploaded))
                            {
                                var teacherWorksheet = excel.Workbook.Worksheets.Single(ws => ws.Name == SAttribute);
                                var cells = teacherWorksheet.Cells;
                                List<List<String>> HeaderCells = new List<List<String>>();
                                j = 0;

                                cn = conE.conectarExpress();
                                cn.Open();
                                cmd_conf = new SqlCommand("select Code,OFile,OSeparator,OField,ODescription,OType,OVar,Okey,isnull(OInitial,0),isnull(OEnd,0),DFile,DField,DType,OFil,DHeader1,DHeader2,DHeader3,DHeader4,DHeader5,DHeader6,isnull(OTrim,0),DLon from [@IC_DFILE] where U_IC_FILE =" + code_file + " and state = 0", cn);
                                rd_conf = cmd_conf.ExecuteReader();

                                sql_campos = null;
                                sql_values = null;
                                sql_where = null;
                                sql_distinct = null;
                                sql = null;
                                coma = null;

                                while (rd_conf.Read())
                                {
                                    if (j > 0 && !String.IsNullOrEmpty(sql_values))
                                    {
                                        coma = ",";
                                    }

                                    if (rd_conf.GetValue(5).ToString() == "F" && rd_conf.GetBoolean(13) == false)
                                    {
                                        try
                                        {
                                            List<String> dHeaderCells = new List<String>();
                                            sql_campos = sql_campos + coma + rd_conf.GetValue(11).ToString();
                                            dHeaderCells.Add(rd_conf.GetValue(3).ToString());
                                            dHeaderCells.Add(rd_conf.GetValue(5).ToString());
                                            dHeaderCells.Add(rd_conf.GetValue(6).ToString());
                                            HeaderCells.Add(dHeaderCells);
                                            sql_values = sql_values + coma + "'" + cells[rd_conf.GetValue(3).ToString() + xHeader].Value.ToString().Replace("'", "''") + "'";
                                        }
                                        catch (Exception)
                                        {
                                            sql_values = sql_values + coma + "''";
                                        }
                                        j++;
                                    }
                                    else if (rd_conf.GetValue(5).ToString() == "V" && rd_conf.GetBoolean(13) == false)
                                    {
                                        List<String> dHeaderCells = new List<String>();
                                        sql_campos = sql_campos + coma + rd_conf.GetValue(11).ToString();
                                        dHeaderCells.Add(rd_conf.GetValue(3).ToString());
                                        dHeaderCells.Add(rd_conf.GetValue(5).ToString());
                                        dHeaderCells.Add(rd_conf.GetValue(6).ToString());
                                        HeaderCells.Add(dHeaderCells);
                                        if (rd_conf.GetValue(6).ToString() == "@HGUI")
                                        {
                                            var valor = hostgroupid;
                                            if (rd_conf.GetValue(12).ToString() == "A")
                                            {
                                                sql_values = sql_values + coma + "'" + valor + "'";
                                            }
                                        }
                                        else if (rd_conf.GetValue(6).ToString() == "@CIA")
                                        {
                                            var valor = cliente;
                                            if (rd_conf.GetValue(12).ToString() == "A")
                                            {
                                                sql_values = sql_values + coma + "'" + valor + "'";
                                            }
                                        }
                                        else
                                        {
                                            //Programar
                                        }
                                        j++;
                                    }

                                    //if (rd_conf.GetValue(5).ToString() == "C" && rd_conf.GetBoolean(13) == false)
                                    //{
                                    //    try
                                    //    {
                                    //        List<String> dHeaderCells = new List<String>();
                                    //        sql_campos = sql_campos + coma + rd_conf.GetValue(11).ToString();
                                    //        dHeaderCells.Add(rd_conf.GetValue(3).ToString());
                                    //        dHeaderCells.Add(rd_conf.GetValue(5).ToString());
                                    //        dHeaderCells.Add(rd_conf.GetValue(6).ToString());
                                    //        HeaderCells.Add(dHeaderCells);
                                    //        sql_values = sql_values + coma + "'" + rd_conf.GetValue(3).ToString() + "'";
                                    //    }
                                    //    catch (Exception)
                                    //    {
                                    //        sql_values = sql_values + coma + "''";
                                    //    }
                                    //    j++;
                                    //}
                                }
                                sql_values_total = "select " + sql_values;
                                StringBuilder sbMaster = new StringBuilder();

                                sbMaster.Append(sql_values_total);
                                int rowCount = cells["A:A"].Count();
                                if (rowCount > xHeader)
                                {

                                    for (int i = xHeader + 1; i <= rowCount; i++)
                                    {
                                        sql_values = string.Empty;
                                        coma = string.Empty;
                                        //sbDetail.Clear();
                                        StringBuilder sbDetail = new StringBuilder();
                                        foreach (var lista in HeaderCells)
                                        {
                                            //if (!String.IsNullOrEmpty(sql_values))
                                            if (!String.IsNullOrEmpty(sbDetail.ToString()))
                                            {
                                                coma = ",";
                                            }
                                            if (lista[1].ToString() == "V" && lista[2].ToString() == "@HGUI")
                                            {
                                                //sb.Append(sql_values);
                                                sbDetail.Append(coma).Append("'").Append(hostgroupid).Append("'");
                                                //sql_values = sql_values + coma + "'" + hostgroupid + "'";
                                            }
                                            else if (lista[1].ToString() == "V" && lista[2].ToString() == "@CIA")
                                            {
                                                //sb.Append(sql_values);
                                                sbDetail.Append(coma).Append("'").Append(cliente).Append("'");
                                                //sql_values = sql_values + coma + "'" + cliente + "'";
                                            }
                                            else
                                            {
                                                //sb.Append(sql_values);
                                                if (cells[lista[0].ToString() + i].Value == null)
                                                {
                                                    sbDetail.Append(coma).Append("'").Append(cells[lista[0].ToString() + i].Value).Append("'");
                                                }
                                                else
                                                {
                                                    sbDetail.Append(coma).Append("'").Append(cells[lista[0].ToString() + i].Value.ToString().Replace("'", "''")).Append("'");
                                                }
                                                //sql_values = sql_values + coma + "'" + cells[lista[0].ToString() + i].Value + "'";
                                            }
                                        }
                                        //sbMaster.Append(" union all select ");
                                        //sbMaster.Append(sbDetail.ToString()).AppendLine();
                                        //sbMaster.Append(sbDetail.ToString() + Environment.NewLine + " ");
                                        //sbMaster.Append(" union all select " + sbDetail.ToString() + Environment.NewLine + " ");
                                        sbMaster.Append(" union all select " + sbDetail);
                                        //sbMaster.Insert(sbMaster.ToString().Length, " union all select " + sbDetail.ToString() + Environment.NewLine + " ");
                                        //sql_values_total = sbMaster.ToString(0,sbMaster.Length);
                                        //sql_values_total = sql_values_total + " union all select " + sql_values;
                                    }
                                }
                                cn.Close();
                                //sql = "insert " + destino_file + " (" + sql_campos + ") " + sql_values_total;
                                sql = "insert " + destino_file + " (" + sql_campos + ") " + Convert.ToString(sbMaster);
                                EjecutarScript(sql);
                            }
                            //using (XLWorkbook wb = new XLWorkbook())
                            //{

                            //}
                            //var fileName = @"C:\ExcelFile.xlsx";
                            //var connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ruta + namefile + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\""; ;
                            //using (var conn = new System.Data.OleDb.OleDbConnection(connectionString))
                            //{
                            //    conn.Open();

                            //    var sheets = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                            //    using (var cmd = conn.CreateCommand())
                            //    {
                            //        cmd.CommandText = "SELECT * FROM [" + sheets.Rows[0]["TABLE_NAME"].ToString() + "] ";

                            //        var adapter = new System.Data.OleDb.OleDbDataAdapter(cmd);
                            //        var ds = new DataSet();
                            //        adapter.Fill(ds);
                            //    }
                            //}
                        }
                        if (!System.IO.File.Exists(historico + namefile))
                        {
                            System.IO.File.Move(ruta + namefile, historico + namefile);
                            //Registrar log
                            RegistraEvento(hostgroupid, cliente, proceso, namefile, destino_file, 1, null, "Envío correcto de interfaz", pk);
                        }
                        else
                        {
                            System.IO.File.Move(ruta + namefile, historico + namefile + "-" + DateTime.Now.ToString("yyyyMMddHHmmssfff"));
                            //Registrar log
                            RegistraEvento(hostgroupid, cliente, proceso, namefile, destino_file, 1, null, "Envío correcto de datos - archivo renombrado", pk);
                        }
                    }
                    catch (Exception e)
                    {
                        if (!System.IO.File.Exists(LogErr + namefile))
                        {
                            System.IO.File.Move(ruta + namefile, LogErr + namefile);
                            //Registrar log
                            RegistraEvento(hostgroupid, cliente, proceso, namefile, destino_file, 0, e.Message.ToString(), "Error en envío de datos: " + sql, pk);
                            Resultado = true;
                        }
                        else
                        {
                            System.IO.File.Move(ruta + namefile, LogErr + namefile + "-" + DateTime.Now.ToString("yyyyMMddHHmmssfff"));
                            //Registrar log
                            RegistraEvento(hostgroupid, cliente, proceso, namefile, destino_file, 0, e.Message.ToString(), "Error en envío de datos - archivo renombrado: " + sql, pk);
                            Resultado = true;
                        }
                    }


                }
            }
            //Condicional de tabla a tabla
            else if (Type == 1)
            {
                //Recorrer tabla

                List<List<string>> listTabletoTable = new List<List<string>>();
                j = 0;

                cn = conE.conectarExpress();
                cn.Open();
                cmd_conf = new SqlCommand("select Code,OFile,OSeparator,OField,ODescription,OType,OVar,Okey,isnull(OInitial,0),isnull(OEnd,0),DFile,DField,DType,OFil,DHeader1,DHeader2,DHeader3,DHeader4,DHeader5,DHeader6,isnull(OTrim,0),DLon from [@IC_DFILE] where U_IC_FILE =" + code_file + " and state = 0", cn);
                rd_conf = cmd_conf.ExecuteReader();

                while (rd_conf.Read())
                {
                    if (j > 0 && !String.IsNullOrEmpty(sql_values))
                    {
                        coma = ",";
                    }
                    if (rd_conf.GetValue(5).ToString() == "V" && bool.Parse(rd_conf.GetValue(13).ToString()) == false)
                    { //Variable
                        sql_campos = sql_campos + coma + rd_conf.GetValue(11).ToString();
                        sql_values = sql_values + coma + ObtenerVariableSQL(rd_conf.GetValue(3).ToString(), hostgroupid, cliente);
                    }
                    else if (rd_conf.GetValue(5).ToString() == "C" && bool.Parse(rd_conf.GetValue(13).ToString()) == false)
                    { //Constante
                        sql_campos = sql_campos + coma + rd_conf.GetValue(11).ToString();
                        sql_values = sql_values + coma + "'" + rd_conf.GetValue(3).ToString() + "'";
                    }
                    else if (rd_conf.GetValue(5).ToString() == "F" && bool.Parse(rd_conf.GetValue(13).ToString()) == false) //Campo
                    {
                        sql_campos = sql_campos + coma + rd_conf.GetValue(11).ToString();
                        if ((int.Parse(rd_conf.GetValue(9).ToString()) >= int.Parse(rd_conf.GetValue(8).ToString())) && (int.Parse(rd_conf.GetValue(9).ToString()) + int.Parse(rd_conf.GetValue(8).ToString())) > 0)
                        {
                            sql_values = sql_values + coma + "substring(" + rd_conf.GetValue(3).ToString()
                                                                    + "," + rd_conf.GetValue(8).ToString() + ","
                                                                    + ((int.Parse(rd_conf.GetValue(9).ToString()) - int.Parse(rd_conf.GetValue(8).ToString())) + 1).ToString() + ")";
                        }
                        else
                        {
                            sql_values = sql_values + coma + rd_conf.GetValue(3).ToString();
                        }
                    }
                    else if (rd_conf.GetValue(5).ToString() == "P" && bool.Parse(rd_conf.GetValue(13).ToString()) == false) //Programa
                    {

                    }
                    if (bool.Parse(rd_conf.GetValue(7).ToString()))
                    {
                        string comillas = null;
                        if (rd_conf.GetValue(6).ToString() == "@HGUI")
                        {
                            var valor = hostgroupid;
                            if (rd_conf.GetValue(12).ToString() == "A")
                            {
                                comillas = "'";
                            }
                            sql_where = sql_where + "and " + rd_conf.GetValue(3).ToString() + " = " + comillas + valor + comillas;
                        }
                    }
                    if (rd_conf.GetValue(5).ToString() == "C" || rd_conf.GetValue(5).ToString() == "F" || rd_conf.GetValue(5).ToString() == "V")
                    {
                        listTabletoTable.Add(new List<string> {
                               rd_conf.GetValue(12).ToString()
                        });
                    }
                    j++;
                }
                cn.Close();

                try
                {
                    cn = conE.conectarExpress();
                    cn.Open();
                    //Recorrer 
                    var sql_WhereSqlConf = ObtenerSQLWhere(code_file, hostgroupid, cliente);
                    var sql_JoinSqlConf = ObtenerSQLJoin(code_file, hostgroupid, cliente);
                    cmd_datos = new SqlCommand("select " + sql_distinct + sql_values + " from " + prefijo + sql_JoinSqlConf + " where 1 = 1 " + sql_where + sql_WhereSqlConf, cn);
                    rd_datos = cmd_datos.ExecuteReader();
                    sql = null;
                    while (rd_datos.Read())
                    {
                        string SQLValue = null;
                        string CComa = null;
                        for (int n = 0; n < rd_datos.FieldCount; n++)
                        {
                            if (n > 0)
                            {
                                CComa = ",";
                            }
                            if (listTabletoTable[n][0].ToString() == "A")
                            {
                                SQLValue = SQLValue + CComa + "'" + rd_datos.GetValue(n).ToString() + "'";
                            }
                            else
                            {
                                SQLValue = SQLValue + CComa + rd_datos.GetValue(n).ToString();
                            }
                        }
                        sql = null;

                        sql = "insert into " + destino_file + "(" + sql_campos + ")" + " values(" + SQLValue + ")";
                        EjecutarScript(sql);
                        z++;
                    }
                    cn.Close();
                    //Registro Log
                    if (z > 0)
                    {
                        RegistraEvento(hostgroupid, cliente, proceso, prefijo, destino_file, 1, null, "Envío correcto de datos", pk);
                    }

                }
                catch (Exception e)
                {
                    RegistraEvento(hostgroupid, cliente, proceso, prefijo, null, 0, e.Message.ToString() + " SQL: " + (String.IsNullOrEmpty(sql) ? "" : sql.ToString()), "Error en envío pase de datos, linea " + z.ToString(), pk);
                    Resultado = true;
                }
            }
            //Tabla -> File
            else if (Type == 3)
            {

                //Recorrer tabla
                //List<List<string>> listTabletoFile = new List<List<string>>();
                List<List<string>> listHeader = new List<List<string>>();

                cn = conE.conectarExpress();
                cn.Open();
                cmd_conf = new SqlCommand("select Code,OFile,OSeparator,OField,ODescription,OType,OVar,Okey,isnull(OInitial,0),isnull(OEnd,0),DFile,DField,DType,OFil,DHeader1,DHeader2,DHeader3,DHeader4,DHeader5,DHeader6,isnull(OTrim,0),DLon from [@IC_DFILE] where U_IC_FILE =" + code_file + " and state = 0", cn);
                rd_conf = cmd_conf.ExecuteReader();

                while (rd_conf.Read())
                {
                    //Obtener lista de Header
                    if (xHeader > 0 && xHeader <= 6)
                    {

                        //for (int h = 1; h <= xHeader;h++) {
                        if (bool.Parse(rd_conf.GetValue(13).ToString()) == false)
                        {
                            listHeader.Add(new List<string> {
                                   rd_conf.GetValue(14).ToString(),
                                   rd_conf.GetValue(15).ToString(),
                                   rd_conf.GetValue(16).ToString(),
                                   rd_conf.GetValue(17).ToString(),
                                   rd_conf.GetValue(18).ToString(),
                                   rd_conf.GetValue(19).ToString()
                            });
                        }
                        //}
                    }
                    string seteo_char_pre = string.Empty;
                    string seteo_char_suf = string.Empty;
                    if (!String.IsNullOrEmpty(rd_conf.GetValue(21).ToString()))
                    {
                        if (rd_conf.GetValue(12).ToString() == "A")
                        {
                            seteo_char_pre = "cast(";
                            seteo_char_suf = " as char(" + rd_conf.GetValue(21).ToString() + "))";
                        }
                        else if (rd_conf.GetValue(12).ToString() == "N")
                        {
                            seteo_char_pre = "replicate(' '," + rd_conf.GetValue(21).ToString() + "- len(";
                            if (rd_conf.GetValue(5).ToString() == "V" && bool.Parse(rd_conf.GetValue(13).ToString()) == false)
                            {
                                seteo_char_suf = ")) + rtrim(ltrim(" + ObtenerVariableSQL(rd_conf.GetValue(3).ToString(), hostgroupid, cliente) + "))";
                            }
                            else if (rd_conf.GetValue(5).ToString() == "C" && bool.Parse(rd_conf.GetValue(13).ToString()) == false)
                            {
                                seteo_char_suf = ")) + '" + rd_conf.GetValue(3).ToString() + "'";
                            }
                            else if (rd_conf.GetValue(5).ToString() == "F" && bool.Parse(rd_conf.GetValue(13).ToString()) == false) //Campo
                            {
                                if ((int.Parse(rd_conf.GetValue(9).ToString()) >= int.Parse(rd_conf.GetValue(8).ToString())) && (int.Parse(rd_conf.GetValue(9).ToString()) + int.Parse(rd_conf.GetValue(8).ToString())) > 0)
                                {
                                    seteo_char_suf = ")) + rtrim(ltrim(substring(" + rd_conf.GetValue(3).ToString()
                                                                            + "," + rd_conf.GetValue(8).ToString() + ","
                                                                            + ((int.Parse(rd_conf.GetValue(9).ToString()) - int.Parse(rd_conf.GetValue(8).ToString())) + 1).ToString() + ")))";
                                }
                                else
                                {
                                    seteo_char_suf = ")) + rtrim(ltrim(" + rd_conf.GetValue(3).ToString() + "))";
                                }
                            }
                            else if (rd_conf.GetValue(5).ToString() == "P" && bool.Parse(rd_conf.GetValue(13).ToString())) //Programa
                            {

                            }
                        }
                    }
                    if (j > 0 && !String.IsNullOrEmpty(sql_values))
                    {
                        coma = ",";
                    }
                    if (rd_conf.GetValue(5).ToString() == "V" && bool.Parse(rd_conf.GetValue(13).ToString()) == false)
                    { //Variable
                        sql_values = sql_values + coma + seteo_char_pre + "rtrim(ltrim(" + ObtenerVariableSQL(rd_conf.GetValue(3).ToString(), hostgroupid, cliente) + "))" + seteo_char_suf;
                    }
                    else if (rd_conf.GetValue(5).ToString() == "C" && bool.Parse(rd_conf.GetValue(13).ToString()) == false)
                    { //Constante
                        sql_values = sql_values + coma + seteo_char_pre + "'" + rd_conf.GetValue(3).ToString() + "'" + seteo_char_suf;
                    }
                    else if (rd_conf.GetValue(5).ToString() == "F" && bool.Parse(rd_conf.GetValue(13).ToString()) == false) //Campo
                    {
                        if ((int.Parse(rd_conf.GetValue(9).ToString()) >= int.Parse(rd_conf.GetValue(8).ToString())) && (int.Parse(rd_conf.GetValue(9).ToString()) + int.Parse(rd_conf.GetValue(8).ToString())) > 0)
                        {
                            sql_values = sql_values + coma + seteo_char_pre + "rtrim(ltrim(substring(" + rd_conf.GetValue(3).ToString()
                                                                    + "," + rd_conf.GetValue(8).ToString() + ","
                                                                    + ((int.Parse(rd_conf.GetValue(9).ToString()) - int.Parse(rd_conf.GetValue(8).ToString())) + 1).ToString() + ")))" + seteo_char_suf;
                        }
                        else
                        {
                            sql_values = sql_values + coma + seteo_char_pre + "rtrim(ltrim(" + rd_conf.GetValue(3).ToString() + "))" + seteo_char_suf;
                        }
                    }
                    else if (rd_conf.GetValue(5).ToString() == "P" && bool.Parse(rd_conf.GetValue(13).ToString())) //Programa
                    {

                    }
                    if (bool.Parse(rd_conf.GetValue(7).ToString()))
                    {
                        string comillas = null;
                        if (rd_conf.GetValue(6).ToString() == "@HGUI")
                        {
                            var valor = hostgroupid;
                            if (rd_conf.GetValue(12).ToString() == "A")
                            {
                                comillas = "'";
                            }
                            sql_where = sql_where + "and " + rd_conf.GetValue(3).ToString() + " = " + comillas + valor + comillas;
                        }
                    }
                    j++;
                }

                if (!String.IsNullOrEmpty(sql_values))
                {
                    //Leer informacion de Header
                    List<Object> listHeaderFinal;
                    listHeaderFinal = new List<object>();

                    object[] hh;
                    object[] firstrow = null;
                    firstrow = new Object[listHeader.Count()];
                    for (int c = 0; c < xHeader; c++)
                    {
                        hh = new Object[listHeader.Count()];

                        for (int q = 0; q < listHeader.Count(); q++)
                        {
                            //sql_union1 = sql_union1 + coma_union + "'" + listHeader[q][0].ToString() + "'";
                            hh[q] = listHeader[q][c].ToString();
                        }
                        if (c == 0)
                        {
                            firstrow = hh;
                        }
                        else
                        {
                            listHeaderFinal.Add(hh);
                        }
                    }

                    //sql_union1 = " union all select " + sql_union1;

                    //
                    cn.Close();

                    //Captura de Log
                    string sql_acumulado = string.Empty;
                    try
                    {

                        cn = conE.conectarExpress();
                        cn.Open();
                        //List<List<String>> Rutina = Query<List<String>>("select " + sql_values + " from " + prefijo + " where 1 = 1 " + sql_where, cn);
                        //Recorrer datos
                        var sql_WhereSqlConf = ObtenerSQLWhere(code_file, hostgroupid, cliente);
                        var sql_JoinSqlConf = ObtenerSQLJoin(code_file, hostgroupid, cliente);
                        sql_acumulado = "select " + sql_distinct + sql_values + " from " + prefijo + sql_JoinSqlConf + " where 1 = 1 " + sql_where + sql_WhereSqlConf;
                        cmd_datos = new SqlCommand(sql_acumulado, cn);
                        List<Object> res;
                        res = new List<Object>();
                        rd_datos = cmd_datos.ExecuteReader();
                        object[] oo;
                        for (int v = 0; v < xHeader - 1; v++)
                        {
                            res.Add(listHeaderFinal[v]);
                        }
                        int col = 0;
                        while (rd_datos.Read())
                        {
                            oo = new Object[rd_datos.FieldCount];
                            for (int inc = 0; inc < rd_datos.FieldCount; inc++)
                            {
                                col = rd_datos.FieldCount;
                                rd_datos.GetValues(oo);
                            }
                            res.Add(oo);
                            z++;
                        }
                        cn.Close();
                        string name_file_exp = null;
                        if ((ext.ToUpper() == "XLS" || ext.ToUpper() == "XLSX") && z > 0)
                        {
                            //Recibir List y convertir en Excel
                            DataTable dt = new DataTable();
                            dt = ToDataTable(res, col, firstrow);
                            name_file_exp = Excel(dt, ruta, destino_file, ext, hostgroupid, SAttribute);
                        }
                        else if (ext.ToUpper() == "TXT")
                        {
                            if (res.Count > 1)
                            {
                                DataTable dt = new DataTable();
                                dt = ToDataTable(res, col, firstrow);
                                name_file_exp = ExportTxt(dt, ruta, destino_file, ext, hostgroupid);
                            }
                        }
                        if (z > 0)
                        {
                            RegistraEvento(hostgroupid, cliente, proceso, prefijo, name_file_exp, 1, sql_acumulado, "Exportación de archivo " + ext.ToUpper(), pk);
                        }
                    }
                    catch (Exception e)
                    {
                        RegistraEvento(hostgroupid, cliente, proceso, prefijo, null, 0, e.Message.ToString(), "Error en envío pase de datos, Linea " + z.ToString() + " SQL:" + sql_acumulado, pk);
                        Resultado = true;
                    }
                }
            }
            //Sql
            else if (Type == 4)
            {
                string sql_ejecutar = string.Empty;
                try
                {
                    var sql_exec = SAttribute;
                    sql_ejecutar = Escanear_Variables(sql_exec, hostgroupid, cliente);
                    int i = EjecutarScriptRetorno(sql_ejecutar);
                    if (i > 0)
                    {
                        RegistraEvento(hostgroupid, cliente, proceso, prefijo, i.ToString(), 1, "Ejecucion de Transacciones", sql_ejecutar, pk);
                    }
                }
                catch (Exception e)
                {
                    RegistraEvento(hostgroupid, cliente, proceso, prefijo, null, 0, e.Message.ToString(), "Error en la ejecucion de Transacciones" + sql_ejecutar, pk);
                    Resultado = true;
                }
            }
            //Sql con parametro XML
            else if (Type == 5) {
                string[] files;
                string[] images;
                string pedido = string.Empty;
                string alias = string.Empty;
                string customer = string.Empty;
                string direccion = string.Empty;
                string city = string.Empty;
                string state = string.Empty;
                string postal = string.Empty;
                string country = string.Empty;
                string comentarios = string.Empty;
                string imagenbase64 = string.Empty;
                bool flag = false;
                
                if (modo == 0)
                {
                    files = Directory.GetFiles(ruta, prefijo + "*." + ext);
                }
                else
                {
                    files = Directory.GetFiles(ruta, prefijo + ext);
                }
                foreach (string names in files)
                {

                    string namefile = Path.GetFileName((names));
                    string nameimagen = string.Empty;
                    Console.WriteLine("Archivo XML: " + namefile);
                    if (Path.GetExtension(namefile).ToUpper() == ".XML")
                    {
                        XmlDocument doc = new XmlDocument();
                        doc.Load(ruta + namefile);
                        ArrayXml arreglo = new ArrayXml();
                        arreglo = GetXMLAsString(doc);
                        //Obtener Pedido

                        foreach (DetalleArrayXml detalle in arreglo.Detalle.FindAll(x => x.nodo == "out:StatusReference"))
                        {
                            if (flag == false)
                            {
                                foreach (var subdetalle in detalle.Dvalor.FindAll(x => x.nodo == "out:ReferenceQualifier"))
                                {
                                    if (subdetalle.DDvalor[0].valor.ToString() == "ON")
                                    {
                                        pedido = detalle.Dvalor[1].DDvalor[0].valor.ToString();
                                        flag = true;
                                        Console.WriteLine("Pedido: " + pedido);
                                        break;
                                    }
                                }
                            }
                            else {
                                break;
                            }
                        }
                        flag = false;
                        //Obtener direccion de entrega
                        foreach (DetalleArrayXml detalle in arreglo.Detalle.FindAll(x => x.nodo == "out:StatusLocation"))
                        {
                            if (flag == false) {
                                foreach (var subdetalle in detalle.Dvalor.FindAll(x => x.nodo == "out:PartyQualifier"))
                                {
                                    if (subdetalle.DDvalor[0].valor.ToString() == "ST")
                                    {
                                        xDetalleArrayXml subdetalle_d = detalle.Dvalor.Find(x => x.nodo == "out:EntityAlias");
                                        if (subdetalle_d != null)
                                        {
                                            alias = subdetalle_d.DDvalor[0].valor.ToString();
                                        }
                                        subdetalle_d = detalle.Dvalor.Find(x => x.nodo == "out:EntityName");
                                        if (subdetalle_d != null) {
                                            customer = subdetalle_d.DDvalor[0].valor.ToString();   
                                        }
                                        subdetalle_d = detalle.Dvalor.Find(x => x.nodo == "out:Address1");
                                        if (subdetalle_d != null)
                                        {
                                            direccion = subdetalle_d.DDvalor[0].valor.ToString();
                                        }
                                        subdetalle_d = detalle.Dvalor.Find(x => x.nodo == "out:City");
                                        if (subdetalle_d != null)
                                        {
                                            city = subdetalle_d.DDvalor[0].valor.ToString();
                                        }
                                        subdetalle_d = detalle.Dvalor.Find(x => x.nodo == "out:State");
                                        if (subdetalle_d != null)
                                        {
                                            state = subdetalle_d.DDvalor[0].valor.ToString();
                                        }
                                        subdetalle_d = detalle.Dvalor.Find(x => x.nodo == "out:PostalCode");
                                        if (subdetalle_d != null)
                                        {
                                            postal = subdetalle_d.DDvalor[0].valor.ToString();
                                        }
                                        subdetalle_d = detalle.Dvalor.Find(x => x.nodo == "out:CountryCode");
                                        if (subdetalle_d != null)
                                        {
                                            country = subdetalle_d.DDvalor[0].valor.ToString();
                                        }
                                        Console.WriteLine("Alias: " + alias);
                                        flag = true;
                                        break;
                                    }
                                }
                            }
                        }
                        flag = false;
                        //Obtener fecha y hora de despacho
                        string departuredate = string.Empty;
                        string departuretime = string.Empty;
                        foreach (DetalleArrayXml detalle in arreglo.Detalle.FindAll(x => x.nodo == "out:StatusLocation"))
                        {
                            if (flag == false)
                            {
                                foreach (xDetalleArrayXml subdetalle in detalle.Dvalor.FindAll(x => x.nodo == "out:ShipmentStatusEntityDateTime"))
                                {
                                    if (subdetalle.DDvalor[0].valor.ToString() == "AAD")
                                    {
                                        string[] departure = subdetalle.DDvalor[1].valor.Split('T');
                                        departuredate = departure[0].ToString();
                                        departuretime = departure[1].ToString();
                                        Console.WriteLine("AAD: " + departuredate + " " + departuretime);
                                        flag = true;
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                        //Obtener comentarios
                        foreach (DetalleArrayXml detalle in arreglo.Detalle.FindAll(x => x.nodo == "out:StatusReference"))
                        {
                            if (flag == false)
                            {
                                foreach (var subdetalle in detalle.Dvalor.FindAll(x => x.nodo == "out:ReferenceQualifier"))
                                {
                                    if (subdetalle.DDvalor[0].valor.ToString() == "OTI")
                                    {
                                        comentarios = detalle.Dvalor[1].DDvalor[0].valor.ToString();
                                        flag = true;
                                        Console.WriteLine("Comentarios: " + comentarios);
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                        flag = false;
                        /*OBTENER JPG*/
                        string name = string.Empty;
                        try
                        {
                            images = Directory.GetFiles(LogErr, "POD_" + pedido + "_*." + "JPEG");
                            foreach (string imagenes in images)
                            {
                                nameimagen = Path.GetFileName((imagenes));
                                if (Path.GetExtension(imagenes).ToUpper() == ".JPEG")
                                {
                                    name = nameimagen;
                                    imagenbase64 = ImageToBase64(LogErr + nameimagen);
                                }
                            }
                        }
                        catch (Exception e) {
                            RegistraEvento(hostgroupid, cliente, proceso, prefijo,pedido.ToString(), 1, "Error al momento de buscar el archivo JPG", "", pk);
                        }
                                /**/
                        string sql_ejecutar = string.Empty;
                        try
                        {
                            var sql_exec = SAttribute;
                            sql_ejecutar = Escanear_Variables_v1(sql_exec, hostgroupid, cliente,pedido,customer,alias,direccion,city,state,postal,country,comentarios,name,imagenbase64, departuredate, departuretime);
                            int i = EjecutarScriptRetorno(sql_ejecutar);
                            if (i > 0)
                            {
                                RegistraEvento(hostgroupid, cliente, proceso, prefijo, pedido.ToString(), 1, "Ejecucion de Transacciones", "", pk);
                                MoveHis(historico, LogErr, namefile, hostgroupid, ruta, cliente, proceso, pedido, nameimagen, destino_file, Resultado,pk,sql);
                            }
                        }
                        catch (Exception e)
                        {
                            RegistraEvento(hostgroupid, cliente, proceso, prefijo, pedido.ToString(), 0, e.Message.ToString(), "Error en la ejecucion de Transacciones" + sql_ejecutar, pk);
                            Resultado = true;
                        }
                    }

                    /**/
                    /**/
                   
                }
            }
                return Resultado;
        }

        /*Rutina para backup - HIS (Schneider)*/
        public bool MoveHis(string historico,string LogErr, string namefile,string hostgroupid,string ruta,string cliente, string proceso, string pedido, string nameimagen,string destino_file,bool Resultado,int pk,string sql) {
            try
            {
                /*XML*/
                string carpeta_anio_mes = string.Empty;
                carpeta_anio_mes = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + @"\";
                if (!System.IO.Directory.Exists(historico + carpeta_anio_mes))
                {
                    System.IO.Directory.CreateDirectory(historico + carpeta_anio_mes);
                }
                if (!System.IO.Directory.Exists(LogErr + @"HIS\" + carpeta_anio_mes))
                {
                    System.IO.Directory.CreateDirectory(LogErr + @"HIS\" + carpeta_anio_mes);
                }


                if (!System.IO.File.Exists(historico + carpeta_anio_mes + namefile))
                {
                    System.IO.File.Move(ruta + namefile, historico + carpeta_anio_mes + namefile);
                    //Registrar log
                    RegistraEvento(hostgroupid, cliente, proceso, namefile, pedido.ToString(), 1, "Pedido: " + pedido.ToString(), "Envío correcto de interfaz", pk);
                }
                else
                {
                    System.IO.File.Move(ruta + namefile, historico + carpeta_anio_mes + namefile + "-" + DateTime.Now.ToString("yyyyMMddHHmmssfff"));
                    //Registrar log
                    RegistraEvento(hostgroupid, cliente, proceso, namefile, pedido.ToString(), 1, "Pedido: " + pedido.ToString(), "Envío correcto de datos - archivo renombrado", pk);
                }
                /*JPG*/
                if ((!System.IO.File.Exists(LogErr + @"HIS\" + carpeta_anio_mes + nameimagen)) && !String.IsNullOrEmpty(nameimagen.ToString()))
                {
                    System.IO.File.Move(LogErr + nameimagen, LogErr + @"HIS\" + carpeta_anio_mes + nameimagen);
                    //Registrar log
                    RegistraEvento(hostgroupid, cliente, proceso, nameimagen, pedido.ToString(), 1, "Pedido: " + pedido.ToString(), "Envío correcto de imagen", pk);
                }
                else if (!String.IsNullOrEmpty(nameimagen.ToString()))
                {
                    System.IO.File.Move(LogErr + nameimagen, LogErr + @"HIS\" + carpeta_anio_mes + nameimagen + "-" + DateTime.Now.ToString("yyyyMMddHHmmssfff"));
                    //Registrar log
                    RegistraEvento(hostgroupid, cliente, proceso, nameimagen, pedido.ToString(), 1, "Pedido: " + pedido.ToString(), "Envío correcto de imagen - archivo renombrado", pk);
                }
            }
            catch (Exception e)
            {
                if (!System.IO.File.Exists(LogErr + namefile))
                {
                    System.IO.File.Move(ruta + namefile, LogErr + namefile);
                    //Registrar log
                    RegistraEvento(hostgroupid, cliente, proceso, namefile, destino_file, 0, e.Message.ToString(), "Error en envío de datos: " + sql, pk);
                    Resultado = true;
                }
                else
                {
                    System.IO.File.Move(ruta + namefile, LogErr + namefile + "-" + DateTime.Now.ToString("yyyyMMddHHmmssfff"));
                    //Registrar log
                    RegistraEvento(hostgroupid, cliente, proceso, namefile, destino_file, 0, e.Message.ToString(), "Error en envío de datos - archivo renombrado: " + sql, pk);
                    Resultado = true;
                }
            }
            return Resultado;
        }

        /*Rutina para Schneider (carga de imagenes)*/
        public static string ImageToBase64(string _imagePath)
        {
            string _base64String = null;

            using (System.Drawing.Image _image = System.Drawing.Image.FromFile(_imagePath))
            {
                using (MemoryStream _mStream = new MemoryStream())
                {
                    _image.Save(_mStream, _image.RawFormat);
                    byte[] _imageBytes = _mStream.ToArray();
                    _base64String = Convert.ToBase64String(_imageBytes);

                    return "data:image/jpg;base64," + _base64String;
                }
            }
        }
        public string Escanear_Variables_v1(string sql_exec, string hostgroupid, string cliente,string delivery, string customer, string alias, string direccion,string city, string state, string postal, string country,string comentarios,string name,string imagen,string departuredate, string departuretime)
        {
            sql_exec = String.IsNullOrEmpty(sql_exec) ? "" : sql_exec;
            int scan = sql_exec.IndexOf("@@");
            string cadena = string.Empty;
            while (scan > 1)
            {
                var scan_1 = sql_exec.IndexOf(" ", scan);
                var variable = sql_exec.Substring(scan, (scan_1 - scan));
                cadena = sql_exec.Replace(@variable, obtener_valor_variable_v1(variable, hostgroupid, cliente, delivery,customer,alias,direccion,city,state,postal,country,comentarios,name,imagen,departuredate,departuretime));
                sql_exec = cadena;
                scan = sql_exec.IndexOf("@@", scan + 1);
            }
            if (String.IsNullOrEmpty(cadena))
            {
                cadena = sql_exec;
            }
            return cadena;
        }

        public string obtener_valor_variable_v1(string variable, string hostgroupid, string cliente, string delivery, string customer, string alias, string direccion, string city, string state, string postal, string country, string comentarios,string name, string imagen,string departuredate,string departuretime)
        {
            if (variable == "@@HGUI")
            {
                return "'" + hostgroupid + "'";
            }
            else if (variable == "@@CIA")
            {
                return "'" + cliente + "'";

            }
            else if (variable == "@@delivery")
            {
                return "'" + delivery + "'";
            }
            else if (variable == "@@Customer")
            {
                return "'" + customer + "'";
            }
            else if (variable == "@@Alias_destination")
            {
                return "'" + alias + "'";
            }
            else if (variable == "@@Address_destination")
            {
                return "'" + direccion + "'";
            }
            else if (variable == "@@City_destination")
            {
                return "'" + city + "'";
            }
            else if (variable == "@@State_destination")
            {
                return "'" + state + "'";
            }
            else if (variable == "@@Postal_Code_destination")
            {
                return "'" + postal + "'";
            }
            else if (variable == "@@Country_destination")
            {
                return "'" + country + "'";
            }
            else if (variable == "@@Comments")
            {
                return "'" + comentarios + "'";
            }
            else if (variable == "@@name")
            {
                return "'" + name + "'";
            }
            else if (variable == "@@imagen")
            {
                return "'" + imagen + "'";
            }
            else if (variable == "@@Departure_date")
            {
                return "'" + departuredate + "'";
            }
            else if (variable == "@@Departure_time")
            {
                return "'" + departuretime + "'";
            }
            else
            {
                return string.Empty;
            }
        }
        /**/
        public string Escanear_Variables(string sql_exec,string hostgroupid,string cliente) {
            sql_exec = String.IsNullOrEmpty(sql_exec) ? "" : sql_exec;
            int scan = sql_exec.IndexOf("@@");
            string cadena = string.Empty;
            while (scan > 1) {
                var scan_1 = sql_exec.IndexOf(" ",scan);
                var variable = sql_exec.Substring(scan, (scan_1 - scan));
                cadena = sql_exec.Replace(@variable, @obtener_valor_variable(variable, hostgroupid,cliente));
                sql_exec = cadena;
                scan = sql_exec.IndexOf("@@",scan + 1);
            }
            if (String.IsNullOrEmpty(cadena)) {
                cadena = sql_exec;        
            }
            return cadena;
        }

        public string obtener_valor_variable (string variable, string hostgroupid,string cliente)
        {
            if (variable == "@@HGUI")
            {
                return "'" + hostgroupid + "'";
            }
            else if (variable == "@@CIA")
            {
                return "'" + cliente + "'";

            }
            else {
                return string.Empty;
            }
        }

        public static DataTable ToDataTable<Object>(List<Object> items,int col,object[] fisrtrow)
        {
            DataTable dataTable = new DataTable();

            for (int q = 0;q< col;q++)

            {
                dataTable.Columns.Add(fisrtrow[q].ToString());
            }

            foreach (object item in items)
            {
                var values = new object[col];
                for (int i = 0; i < col; i++)
                {
                    values[i] = ((object[])item)[i].ToString();
                }
                dataTable.Rows.Add(values);
            }
            return dataTable;
            //DataTable dataTable = new DataTable(typeof(T).Name);

            ////Get all the properties
            //PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            //foreach (PropertyInfo prop in Props)
            //{
            //    //Defining type of data column gives proper data table
            //    var type = (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(prop.PropertyType) : prop.PropertyType);
            //    //Setting column names as Property names
            //    dataTable.Columns.Add(prop.Name, type);
            //}
            //foreach (T item in items)
            //{
            //    var values = new object[Props.Length];
            //    for (int i = 0; i < Props.Length; i++)
            //    {
            //        //inserting property values to datatable rows
            //        values[i] = Props[i].GetValue(item, null);
            //    }
            //    dataTable.Rows.Add(values);
            //}
            //put a breakpoint here and check datatable

        }

        public string ExportTxt(DataTable dt,string ruta,string destino_file,string ext,string hostgroupid) {
            StreamWriter swExtLogFile = new StreamWriter(ruta + destino_file + "" + hostgroupid + "." + ext, true);
            int i;
            //swExtLogFile.Write(Environment.NewLine);
            foreach (DataRow row in dt.Rows)
            {
                object[] array = row.ItemArray;
                for (i = 0; i < array.Length - 1; i++)
                {
                    swExtLogFile.Write(array[i].ToString());
                }
                if (array[0].ToString().Trim() != "")
                {
                    swExtLogFile.WriteLine(array[i].ToString());
                }
            }
            swExtLogFile.Flush();
            swExtLogFile.Close();
            return destino_file + "" + hostgroupid + "." + ext;
        }

        public string Excel(DataTable dt,string ruta,string destino_file, string ext,string hostgroupid,string SAttribute)
        {

            string folderPath = ruta;
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    foreach (System.Data.DataColumn col in dt.Columns) col.ReadOnly = false;

                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        dt.Rows[i][j] = dt.Rows[i][j].ToString().Replace(Convert.ToChar((byte)0x1F), ' '); ;
                    }
                }
                if (String.IsNullOrEmpty(SAttribute))
                {
                    SAttribute = destino_file;
                }

                //wb.Worksheets.Add(dt, SAttribute);
                var worksheet = wb.AddWorksheet(dt, SAttribute);
                worksheet.Tables.FirstOrDefault().ShowAutoFilter = false;

                //worksheet.Row(1).Delete();
                //worksheet.FirstRow().Delete();

                //worksheet.FirstRow();
                wb.SaveAs(folderPath + destino_file + "_" + hostgroupid + "." + ext);
                return destino_file + "_" + hostgroupid + "." + ext;
            }

            //using (XLWorkbook wb = new XLWorkbook())
            //{

            //    //var dt = integracion.GetDataSource(ConfigRpt);

            //    for (int i = 0; i < dt.Rows.Count; i++)
            //    {
            //        foreach (System.Data.DataColumn col in dt.Columns) col.ReadOnly = false;

            //        for (int j = 0; j < dt.Columns.Count; j++)
            //        {
            //            dt.Rows[i][j] = dt.Rows[i][j].ToString().Replace(Convert.ToChar((byte)0x1F), ' '); ;
            //        }
            //    }

            //    dt.TableName = "ExcelReport";

            //    wb.Worksheets.Add(dt);
            //    //wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //    wb.Style.Font.Bold = true;

            //    //Response.Clear();
            //    //Response.Buffer = true;
            //    //Response.Charset = "";
            //    //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //    //Response.AddHeader("content-disposition", "attachment;filename= ExcelReport.xlsx");


            //    using (MemoryStream MyMemoryStream = new MemoryStream())
            //    {
            //        wb.SaveAs(MyMemoryStream);
            //        //MyMemoryStream.WriteTo(Response.OutputStream);
            //        //Response.Flush();
            //        //Response.End();
            //    }
            //}
        }
        //public List<T> Query<T>(string query, SqlConnection cn) where T : new()
        //{
        //    List<T> res = new List<T>();
        //    SqlCommand q = new SqlCommand(query, cn);
        //    SqlDataReader r = q.ExecuteReader();
        //    while (r.Read())
        //    {
        //        T t = new T();

        //        for (int inc = 0; inc < r.FieldCount; inc++)
        //        {


        //            //Type type = t.GetType();
        //            //PropertyInfo prop = type.GetProperty(r.GetName(inc));

        //            //prop.SetValue(t, r.GetValue(inc), null);
        //        }

        //        res.Add(t);
        //    }
        //    r.Close();

        //    return res;

        //}

        private string ObtenerVariableSQL(string variable,string hostgroupid, string cliente) {
            SqlConnection cn_sql = null;
            SqlCommand cmd = null;
            SqlDataReader rd = null;
            ConexionExpress con = new ConexionExpress();
            string CadenaSql = null;
            cn_sql = con.conectarExpress();
            cn_sql.Open();
            cmd = new SqlCommand("select VarSQL from [@IC_VARSQL] where Variable = '" + variable + "'", cn_sql);
            rd = cmd.ExecuteReader();
            while (rd.Read()) {
                CadenaSql = rd.GetValue(0).ToString();
            }
            cn_sql.Close();
            CadenaSql = Escanear_Variables(CadenaSql,hostgroupid,cliente);
            return CadenaSql;
        }

        private string ObtenerSQLJoin(int file,string hostgroupid,string cliente)
        {
            SqlConnection cn_sql = null;
            SqlCommand cmd = null;
            SqlDataReader rd = null;
            ConexionExpress con = new ConexionExpress();
            string CadenaSql = null;
            cn_sql = con.conectarExpress();
            cn_sql.Open();
            cmd = new SqlCommand("select VarSQL from [@IC_JOINSQL] where U_EX_FILE = " + file, cn_sql);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                CadenaSql = rd.GetValue(0).ToString();
            }
            cn_sql.Close();
            CadenaSql = Escanear_Variables(CadenaSql, hostgroupid, cliente);
            return CadenaSql;
        }

        private string ObtenerSQLWhere(int file,string hostgroupid,string cliente)
        {
            SqlConnection cn_sql = null;
            SqlCommand cmd = null;
            SqlDataReader rd = null;
            ConexionExpress con = new ConexionExpress();
            string CadenaSql = null;
            cn_sql = con.conectarExpress();
            cn_sql.Open();
            cmd = new SqlCommand("select VarSQL from [@IC_WHERESQL] where U_EX_FILE = " + file, cn_sql);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                CadenaSql = rd.GetValue(0).ToString();
            }
            cn_sql.Close();
            CadenaSql = Escanear_Variables(CadenaSql, hostgroupid, cliente);
            return CadenaSql;
        }

        private void EjecutarScript(string strSql)
        {
            //try
            //{
                SqlConnection cn_sql = null;
                SqlCommand cmd = null;
                ConexionExpress con = new ConexionExpress();
                cn_sql = con.conectarExpress();
                cn_sql.Open();
                cmd = new SqlCommand(strSql, cn_sql);
                cmd.CommandTimeout = 1000;
                cmd.ExecuteNonQuery();
                cn_sql.Close();
                //return null;
            //}
            //catch (Exception e)
            //{
            //    //System.Console.WriteLine(e);
            //    return e.Message.ToString();
            //}
        }

        private int EjecutarScriptRetorno(string strSql)
        {
            //try
            //{
            SqlConnection cn_sql = null;
            SqlCommand cmd = null;
            ConexionExpress con = new ConexionExpress();
            cn_sql = con.conectarExpress();
            cn_sql.Open();
            cmd = new SqlCommand(strSql, cn_sql);
            cmd.CommandTimeout = 1000;
            int i = cmd.ExecuteNonQuery();
            cn_sql.Close();
            return i;
            //return null;
            //}
            //catch (Exception e)
            //{
            //    //System.Console.WriteLine(e);
            //    return e.Message.ToString();
            //}
        }

        //public static void uploadfilesftp(string origen, string host, int port, string username, string password, string workingdirectory)
        //{

        //    //FileInfo f = new FileInfo(@"c:SAP_INV_20151206171543.txt");
        //    FileInfo f = new FileInfo(origen);
        //    String uploadfile = f.FullName;

        //    using (var sftp = new SftpClient(host, port, username, password))
        //    {
        //        sftp.Connect();
        //        sftp.ChangeDirectory(workingdirectory);
        //        using (var filestream = new FileStream(uploadfile, FileMode.Open))
        //        {
        //            sftp.BufferSize = 4 * 1024;
        //            sftp.UploadFile(filestream, Path.GetFileName(uploadfile), null);
        //        }

        //        sftp.Disconnect();
        //        sftp.Dispose();
        //    }

        //}
        public static void Downloadfileftp() {
            //FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://" + "190.119.250.147/IN/SOURCE/ITFPRV_20170713130302.txt");
            //request.Method = WebRequestMethods.Ftp.DownloadFile;
            //request.Credentials = new NetworkCredential("ClientePrueba", "cevalogistics");
            //FtpWebResponse response = (FtpWebResponse)request.GetResponse();

            //Stream responseStream = response.GetResponseStream();
            //StreamReader reader = new StreamReader(responseStream);
            //Console.WriteLine(reader.ReadToEnd());

            //Console.WriteLine("Download Complete, status {0}", response.StatusDescription);

            //reader.Close();
            //response.Close();

            FtpWebRequest dirFtp = ((FtpWebRequest)FtpWebRequest.Create("ftp://" + "190.119.250.147/IN/SOURCE/"));

            // Los datos del usuario (credenciales)
            NetworkCredential cr = new NetworkCredential("ClientePrueba", "cevalogistics");
            dirFtp.Credentials = cr;

            // El comando a ejecutar
            dirFtp.Method = "LIST";

            // También usando la enumeración de WebRequestMethods.Ftp
            dirFtp.Method = WebRequestMethods.Ftp.ListDirectoryDetails;

            // Obtener el resultado del comando
            StreamReader reader =
                new StreamReader(dirFtp.GetResponse().GetResponseStream());

            // Leer el stream
            string res = reader.ReadToEnd();

            // Mostrarlo.
            Console.WriteLine(res);

            // Cerrar el stream abierto.
            reader.Close();
        }

        //public static void uploadfileftp(string origen, string host, int port, string username, string password, string workingdirectory)
        //{
        //    //http://geeks.ms/blogs/dcerredelo/archive/2012/05/02/c-a-fondo-las-clases-system-net-ftpwebrequest-y-system-net-ftpwebresponse.aspx

        //    //string uploads = @"c:\SAP_INV_20151206171543.txt";
        //    if (String.IsNullOrEmpty(host) == false)
        //    {
        //        string uploads = origen;
        //        FileInfo fileinf = new FileInfo(origen);

        //        FtpWebRequest request = FtpWebRequest.Create("ftp://" + host + ":" + port + workingdirectory + fileinf.Name.ToString()) as FtpWebRequest;

        //        request.Credentials = new NetworkCredential(username, password);

        //        request.UsePassive = false;
        //        request.UseBinary = true;
        //        request.KeepAlive = true;
        //        request.Method = WebRequestMethods.Ftp.UploadFile;
        //        request.ContentLength = fileinf.Length;
        //        request.EnableSsl = false;
        //        request.Proxy = null;

        //        int buffLength = 4048;
        //        byte[] buff = new byte[buffLength];
        //        int contentLen;
        //        FileStream fs = fileinf.OpenRead();

        //        try
        //        {
        //            Stream strm = request.GetRequestStream();

        //            contentLen = fs.Read(buff, 0, buffLength);
        //            while (contentLen != 0)
        //            {
        //                strm.Write(buff, 0, contentLen);
        //                contentLen = fs.Read(buff, 0, buffLength);
        //            }
        //            strm.Close();
        //            fs.Close();
        //        }
        //        catch (Exception ex)
        //        {
        //            //MessageBox.Show(ex.Message.ToString());
        //            //Inscribir error
        //        }
        //    }
        //}

        public void RegistraEvento(string HostGroupId, string cliente, string processName, string fileName, string referencia, int status, string messageSystem, string message,int FK_Ruta)
        {
            try
            {      
                ConexionExpress conE = new ConexionExpress();
                using (SqlConnection conn = conE.conectarExpress())
                using (SqlCommand cmd = new SqlCommand("Sp_registareventoceva", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add("@HostGroupId", SqlDbType.VarChar).Value = HostGroupId == null ? string.Empty : HostGroupId; ;
                    cmd.Parameters.Add("@Cliente", SqlDbType.NVarChar).Value = cliente;
                    cmd.Parameters.Add("@ProcessName", SqlDbType.NVarChar).Value = processName;
                    cmd.Parameters.Add("@FileName", SqlDbType.VarChar).Value = fileName == null ? System.Data.SqlTypes.SqlString.Null : fileName;
                    cmd.Parameters.Add("@Reference", SqlDbType.VarChar).Value = referencia == null ? string.Empty : referencia;
                    cmd.Parameters.Add("@Status", SqlDbType.Int).Value = status;
                    cmd.Parameters.Add("@MessageSystem", SqlDbType.VarChar).Value = messageSystem == null ? System.Data.SqlTypes.SqlString.Null : messageSystem;
                    cmd.Parameters.Add("@Message", SqlDbType.VarChar).Value = message == null ? System.Data.SqlTypes.SqlString.Null : message;
                    cmd.Parameters.Add("@FK_Ruta", SqlDbType.Int).Value = FK_Ruta;
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.ToString());
            }

        }
    }
}
