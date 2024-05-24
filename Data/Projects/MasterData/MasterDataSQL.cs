using System;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;
using System.Data;
using Newtonsoft.Json.Linq;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using WinSCP;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;

namespace DataBotV5.Data.Projects.MasterData
{
    /// <summary>
    /// Clase Data encargada de manejo de datos maestros.
    /// </summary>
    class MasterDataSQL
    {
        Credentials.Credentials cred = new Credentials.Credentials();
        ConsoleFormat console = new ConsoleFormat();
        Rooting root = new Rooting();
        CRUD crud = new CRUD();
        Database.Database db2 = new Database.Database();
        ProcessInteraction proc = new ProcessInteraction();
        /// <summary>Método para obtener la gestión.</summary>
        public string GetManagement(string type)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            root.requestDetails = "";
            string respuesta = "";
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB
                sql_select = "select * from gestiones_dm where ESTADO = 'EN PROCESO' and TIPO_DATO = '" + type + "'"; //and METODO = 'MASIVO' 
                //mytable = crud.Select("Databot", sql_select, "automation");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    console.WriteLine("Extraer datos...");
                    root.IdGestionDM = mytable.Rows[0][1].ToString(); //ID GESTION
                    root.BDUserCreatedBy = mytable.Rows[0][2].ToString().ToLower() + "@gbm.net"; //EMPLEADO
                    root.aprobadorDM = mytable.Rows[0][3].ToString(); //APROBADOR
                    root.factorDM = mytable.Rows[0][4].ToString(); //FACTOR
                    root.datagDM = mytable.Rows[0][5].ToString(); //DG
                    root.fechaDM = mytable.Rows[0][10].ToString(); //TS_CREACION
                    root.tipo_gestion = mytable.Rows[0][14].ToString(); //TIPO_GESTION
                    root.metodoDM = mytable.Rows[0][16].ToString();  //METODO
                    root.Subject = "Formulario Creación de " + type.ToLower() + " - Notificación de Finalización de Gestión - #" + root.IdGestionDM;
                    root.requestDetails = mytable.Rows[0][9].ToString();  //GESTION
                    string docaprob = mytable.Rows[0][17].ToString(); //DOC_APROB  root.dm_files_list
                    if (!String.IsNullOrEmpty(docaprob) && docaprob != "[]")
                    {
                        root.doc_aprob = JArray.Parse(docaprob);
                    }
                    string mass_aprob = mytable.Rows[0][18].ToString();  //MAS_APROB
                    if (root.metodoDM == "MASIVO")
                    {
                        console.WriteLine("Buscando el archivo correcto...");
                        try
                        {
                            JArray gestiones;
                            if (!String.IsNullOrEmpty(mass_aprob) && mass_aprob != "[{}]")
                            {
                                gestiones = JArray.Parse(mass_aprob);
                            }
                            else
                            {
                                gestiones = JArray.Parse(root.requestDetails);
                            }

                            for (int i = 0; i < gestiones.Count; i++)
                            {
                                JObject fila = JObject.Parse(gestiones[i].ToString());
                                string adjunto = "";
                                if (!String.IsNullOrEmpty(mass_aprob) && mass_aprob != "[{}]")
                                {
                                    adjunto = fila["APROB"].Value<string>();
                                }
                                else
                                {
                                    adjunto = fila["PLANTILLA"].Value<string>();
                                }
                                string extArchivo = Path.GetExtension(adjunto);
                                if (extArchivo.Substring(0, 4) == ".xls")
                                {
                                    string local_ruta = root.FilesDownloadPath + "\\" + adjunto;
                                    int index = 1;

                                    #region descargar archivo del FTP   
                                    TransferOperationResult transferResult;
                                    TransferOptions transferOptions = new TransferOptions();

                                    SessionOptions sessionOptions = db2.ConnectFTP(1, "databot.gbm.net", 21, "gbmadmin", cred.password_server_web, false, "");

                                    sessionOptions.AddRawSettings("ProxyPort", "0");

                                    using (Session session = new Session())
                                    {
                                        console.WriteLine(" Estableciendo conexion");
                                        session.Open(sessionOptions);
                                        console.WriteLine(" Descargando archivo");
                                        transferOptions.TransferMode = TransferMode.Binary;
                                        string ftp_ruta = "/dm_gestiones_mass/" + root.IdGestionDM + "/" + adjunto;

                                        transferResult = session.GetFiles(ftp_ruta, local_ruta, false, transferOptions);
                                        transferResult.Check();
                                        session.Dispose();
                                    }
                                    #endregion

                                    #region Abre el excel y verifica que sea la plantilla
                                    Excel.Application xlApp;
                                    Excel.Workbook xlWorkBook;
                                    Excel.Worksheet xlWorkSheet;

                                    xlApp = new Excel.Application();
                                    xlApp.Visible = false;
                                    xlApp.DisplayAlerts = false;

                                    xlWorkBook = xlApp.Workbooks.Open(local_ruta);
                                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];

                                    string validacion = xlWorkSheet.Cells[1, index].text.ToString().Trim();
                                    while (validacion != "")
                                    {
                                        if (validacion != "")
                                        {
                                            if (validacion.Substring(0, 1).ToLower() == "x")
                                            {
                                                root.ExcelFile = adjunto;
                                                break;
                                            }
                                        }
                                        index++;
                                        validacion = xlWorkSheet.Cells[1, index].text.ToString().Trim();
                                    }
                                    xlWorkBook.Close();
                                    xlApp.Workbooks.Close();
                                    xlApp.Quit();
                                    proc.KillProcess("EXCEL", true);

                                    #endregion
                                }

                            }
                        }
                        catch (Exception ex)
                        {
                            return "ERROR";
                        }

                    }

                    respuesta = "OK";
                }


            }
            catch (Exception ex)
            {
                respuesta = "ERROR";
            }
            return respuesta;
        }

        /// <summary>Método para cambiar estado en Datos Maestros.</summary>
        public bool ChangeStateDM(string idGestion, string response, string state, DateTime DateCreated)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_update = "";
            string sql_update2 = "";
            string fechaf = DateCreated.ToString("yyyy-MM-dd HH:mm:ss");
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB   
                sql_select = "select * from gestiones_dm where ID_GESTION = " + idGestion;
                //mytable = crud.Select("Databot", sql_select, "automation");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    sql_update = "Update gestiones_dm set ESTADO = '" + state + "',RESPUESTA = '" + response + "',TS_FINALIZACION = '" + fechaf + "' where ID_GESTION = " + idGestion;
                    //crud.Update("Databot", sql_update, "automation");

                    sql_update2 = "INSERT INTO log_gestiones (`ID_GESTION`, `ESTADO`, `RESPUESTA`, `APROBADOR`, `FECHA`) VALUES ('" + idGestion + "','" + state + "','" + response + "','" + "RPAUSER" + "','" + fechaf + "')";

                    try
                    {
                        //crud.Update("Databot", sql_update2, "automation");
                    }
                    catch (Exception) { }
                }
                else
                {

                }

            }
            catch (Exception ex)
            {

            }
            return false;
        }

        /// <summary>Método para agregar un proveedor.</summary>
        public bool AddVendor(string idVendor, string applicant, string idManagement)
        {
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB   
                string select = "select * from proveedores_dm where id_prov = " + idVendor;
                //mytable = crud.Select("Databot", select, "auditoria_bs");
                #endregion

                if (mytable.Rows.Count <= 0)
                {
                    string insert = "INSERT INTO `proveedores_dm`(`id_prov`, `solicitante`, `id_gestion`) VALUES (" + idVendor + ",'" + applicant + "'," + idManagement + ")";
                    //crud.Insert("Databot", insert, "auditoria_bs");
                }
            }
            catch (Exception)
            {

            }
            return false;
        }

        /// <summary>Método para obtener la descripción de un proveedor.</summary>
        public string GetVendorDescription(string VendorCat)
        {
            string vendor_descrip = "";
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB   
                string select = "SELECT * FROM `vendor_description` WHERE `vendor_cat` = '" + VendorCat + "'";
                //mytable = crud.Select("Databot", select, "automation");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    vendor_descrip = mytable.Rows[0]["vendor_descrip"].ToString(); //ID GESTION
                }
                else
                {
                    vendor_descrip = "D_Mix";
                }
            }
            catch (Exception)
            {
                vendor_descrip = "D_Mix";
            }
            return vendor_descrip;

        }

        /// <summary>Método para verificar la versión.</summary>
        public bool CheckVersion(string template, string atachment)
        {
            //leer la tabla, tomar la version
            string ver_file = "";
            //DataTable xx = crud.Select("Databot", "SELECT * FROM `plantilla_versiones` WHERE `Name` LIKE '" + template + "'", "automation");
            //string ver_db = xx.Rows[0]["version"].ToString();

            //leer la version del file
            Shell32.Shell shell = new Shell32.Shell();
            Shell32.Folder objFolder = shell.NameSpace(Path.GetDirectoryName(atachment));

            foreach (Shell32.FolderItem2 item in objFolder.Items())
            {
                if (item.Name == Path.GetFileName(atachment))
                {
                    ver_file = objFolder.GetDetailsOf(item, 18); //18 es etiqueta
                }
            }

            //if (ver_db == ver_file)
            //    return true;
            //else
            return false;
        }
    }
}
