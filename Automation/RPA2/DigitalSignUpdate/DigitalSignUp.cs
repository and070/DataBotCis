using DataBotV5.App.Global;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.MicrosoftTools;
using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace DataBotV5.Automation.RPA2.DigitalSignUpdate
{
    public class DigitalSignUp
    {
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        SapVariants sap = new SapVariants();
        MsExcel excel = new MsExcel();
        Rooting root = new Rooting();
        CRUD crud = new CRUD();
        string respFinal = "";
        Log log = new Log();
        string mandante = "ERP";


        /// <summary>
        /// Robot que actualiza la información de los colaboradores en digital_sign de Smart and Simple a traves de SAP
        /// </summary>
        public void Main()
        {
            DataTable excelDt = crud.Select("SELECT * FROM `digital_sign`", "MIS");
            console.WriteLine("Processing...");
            ///Procesar------------------
            Process(excelDt);
            ///--------------------------
            ///

            if (excelDt.Rows.Count > 0)
            {
                console.WriteLine("Creando estadísticas...");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }



        }
        /// <summary>
        /// Cargar el item de cada sales order en la VF44 para confirmarlo
        /// </summary>
        /// <param name="ExcelFile">el excel que envía el país por email outlook</param>
        private void Process(DataTable ExcelFile)
        {
            #region private variables
            root.ExcelFile = "digitalSign.xlsx";

            //variable de respuesta en caso de que se necesite
            string response = "Se adjunta el reporte de cambios en los colaboradores de digital_sign";
            //DataTable furuto excel de respuesta
            DataTable dtResponse = new DataTable();
            dtResponse = ExcelFile.Clone();

            //variable nombre de la hoja de resultados
            string dtResponseSheetName = "Results";
            //variable nombre del libro de resultados + extension
            string dtResponseBookName = $"ResultsBook{DateTime.Now.ToString("yyyyMMdd")}" + root.ExcelFile;
            //ruta + nombre donde se guardará el excel de resultado
            string dtResponseRoute = root.FilesDownloadPath + "\\" + dtResponseBookName;
            //PLantilla en html para el envío de email
            string htmlEmail = Properties.Resources.emailtemplate1;
            //variable titulo del cuerpo del correo
            string htmlSubject = "Usuarios actualizados en digital_sign";
            //variable contenido del correo: texto, cuadros, tablas, imagenes, etc
            string htmlContents = "";

            //variable remitente del email de respuesta
            string sender = "fvillalobos@gbm.net";
            //variable copias del email de respuesta
            string[] cc = new string[] { "dmeza@gbm.net" };
            //variable ruta de adjunto
            string[] attachments = new string[] { dtResponseRoute };
            //agrega la columna de resultado en el Excel
            dtResponse.Columns.Add("Response");
            dtResponse.Columns.Add("SqlQuery");
            #endregion

            #region loop each excel row

            console.WriteLine("Foreach Excel row...");
            foreach (DataRow row in ExcelFile.Rows)
            {
                #region robot Process
                console.WriteLine(DateTime.Now + " > > > " + "Corriendo RFC de SAP: " + root.BDProcess);
                try
                {
                    #region Parametros de SAP
                    Dictionary<string, string> parametros = new Dictionary<string, string>
                    {
                        ["USUARIO"] = row["user"].ToString()
                    };


                    IRfcFunction func = sap.ExecuteRFC(mandante, "ZFD_GET_USER_DETAILS", parametros);
                    #endregion

                    #region Procesar Salidas del FM
                    string resp = func.GetValue("RESPUESTA").ToString();

                    if (resp == "OK")
                    {
                        bool changeInfo = false;

                        if (row["name"].ToString() != func.GetValue("NOMBRE").ToString()) { changeInfo = true; }
                        if (row["UserID"].ToString() != func.GetValue("IDCOLABC").ToString().TrimStart('0')) { changeInfo = true; }
                        if (row["department"].ToString() != func.GetValue("DEPARTAMENTO").ToString()) { changeInfo = true; }
                        if (row["manager"].ToString() != func.GetValue("SUPERVISOR").ToString()) { changeInfo = true; }
                        if (row["country"].ToString() != func.GetValue("PAIS").ToString()) { changeInfo = true; }
                        if (row["subDivision"].ToString() != func.GetValue("SUB_DIVISION").ToString()) { changeInfo = true; }
                        if (row["email"].ToString() != func.GetValue("EMAIL").ToString()) { changeInfo = true; }
                        if (row["position"].ToString() != func.GetValue("POSICION").ToString()) { changeInfo = true; }



                        DateTime dtS = DateTime.ParseExact(func.GetValue("INGRESO").ToString(), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                        DateTime dtE = DateTime.ParseExact(func.GetValue("SALIDA").ToString(), "yyyy-MM-dd", CultureInfo.InvariantCulture);

                        DateTime dtSOriginal = DateTime.Parse(row["startDate"].ToString());
                        DateTime dtEOriginal = DateTime.Parse(row["endDate"].ToString());

                        if (dtSOriginal != dtS) { changeInfo = true; }
                        if (dtEOriginal != dtE) { changeInfo = true; }

                        if (changeInfo)
                        {
                            DataRow rRow = dtResponse.Rows.Add();
                            rRow["id"] = row["id"].ToString();
                            rRow["user"] = row["user"].ToString();


                            if (row["name"].ToString() != func.GetValue("NOMBRE").ToString()) { rRow["name"] = func.GetValue("NOMBRE").ToString(); }
                            if (row["UserID"].ToString() != func.GetValue("IDCOLABC").ToString().TrimStart('0'))
                            {
                                rRow["UserID"] = func.GetValue("IDCOLABC").ToString().TrimStart('0');
                            }
                            if (row["department"].ToString() != func.GetValue("DEPARTAMENTO").ToString()) { rRow["department"] = func.GetValue("DEPARTAMENTO").ToString(); }
                            if (row["manager"].ToString() != func.GetValue("SUPERVISOR").ToString()) { rRow["manager"] = func.GetValue("SUPERVISOR").ToString(); }
                            if (row["country"].ToString() != func.GetValue("PAIS").ToString()) { rRow["country"] = func.GetValue("PAIS").ToString(); }
                            if (row["subDivision"].ToString() != func.GetValue("SUB_DIVISION").ToString()) { rRow["subDivision"] = func.GetValue("SUB_DIVISION").ToString(); }
                            if (row["email"].ToString() != func.GetValue("EMAIL").ToString()) { rRow["email"] = func.GetValue("EMAIL").ToString(); }
                            if (row["position"].ToString() != func.GetValue("POSICION").ToString()) { rRow["position"] = func.GetValue("POSICION").ToString(); }
                            if (dtSOriginal != dtS) { rRow["startDate"] = dtS; }
                            if (dtEOriginal != dtE) { rRow["endDate"] = dtE; }

                            rRow["Response"] = resp;
                            string query = sqlCreate(rRow);
                            rRow["SqlQuery"] = query;

                            //update en smart and simple
                            if (!string.IsNullOrWhiteSpace(query))
                            {
                                console.WriteLine("Actualizando en S&S....");
                                bool up = crud.Update(query, "MIS");
                                rRow["Response"] = up.ToString();

                                string respo = $"Actualizando colaborador {row["UserID"].ToString()} en Smart & Simple.";
                                log.LogDeCambios("Actualización", root.BDProcess,  root.BDUserCreatedBy , "Actualizar colaborador", respo, "");
                                respFinal = respFinal + "\\n" + respo;

                            }
                        }


                    }
                    else
                    {
                        DataRow rRow = dtResponse.Rows.Add();
                        rRow["id"] = row["id"].ToString();
                        rRow["user"] = row["user"].ToString();
                        rRow["Response"] = "INACTIVO";
                        rRow["token"] = "";
                        rRow["active"] = row["active"].ToString();
                        rRow["createdAt"] = row["createdAt"].ToString();
                        rRow["updatedAt"] = row["updatedAt"].ToString();
                    }



                    #endregion
                }
                catch (Exception ex)
                {
                    console.WriteLine(DateTime.Now + " > > > " + " Finishing process " + ex.Message);
                    DataRow rRow = dtResponse.Rows.Add();
                    rRow["id"] = row["id"].ToString();
                    rRow["user"] = row["user"].ToString();
                    rRow["Response"] = ex.Message;
                    rRow["token"] = "";
                    rRow["active"] = row["active"].ToString();
                    rRow["createdAt"] = row["createdAt"].ToString();
                    rRow["updatedAt"] = row["updatedAt"].ToString();
                }

                #endregion
            }
            dtResponse.AcceptChanges();
            #endregion
            #region Create results Excel
            console.WriteLine("Save Excel...");
            excel.CreateExcel(dtResponse, dtResponseSheetName, dtResponseRoute);
            #endregion
            #region SendEmail
            console.WriteLine("Send Email...");
            //se agregan los parametros anterior al html del cuerpo del email de respuesta
            htmlEmail = htmlEmail.Replace("{subject}", htmlSubject).Replace("{cuerpo}", response).Replace("{contenido}", htmlContents);
            mail.SendHTMLMail(htmlEmail, new string[] { sender }, $"Actualización de usuarios en digital_sign {DateTime.Now.ToString("dd/MM/yyyy")}", cc, attachments);

            #endregion

            root.requestDetails = respFinal;
            root.BDUserCreatedBy = "FVILLALOBOS";

        }

        private string sqlCreate(DataRow excelRow)
        {
            console.WriteLine("Creando Query");
            string sqlQuery = "UPDATE `digital_sign` SET ";
            try
            {

                string id = excelRow["id"].ToString();
                string sqlQuery2 = $" WHERE id = {id};";

                foreach (DataColumn column in excelRow.Table.Columns)
                {
                    string columna = column.ToString();
                    if (columna != "id" && columna != "token" && columna != "user" && columna != "SqlQuery" && columna != "Response" && columna != "active" && columna != "createdAt" && columna != "updatedAt")
                    {
                        string value = excelRow[columna].ToString();

                        if (value != "")
                        {
                            if (columna == "endDate" || columna == "startDate")
                            {
                                value = DateTime.Parse(value).ToString("yyyy-MM-dd HH:mm:ss");
                            }
                            sqlQuery = sqlQuery + columna + " = '" + value + "', ";
                        }
                    }
                }
                if (sqlQuery != "UPDATE `digital_sign` SET ")
                {
                    //significa que si hubo un cambio realmente
                    sqlQuery = sqlQuery.Substring(0, sqlQuery.Length - 2);
                    sqlQuery = sqlQuery + $", token = NULL, updatedAt = '{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}'";
                    sqlQuery = sqlQuery + sqlQuery2;

                }
                else
                {
                    sqlQuery = "";
                }



            }
            catch (Exception ex)
            {
                sqlQuery = ex.ToString();
            }

            return sqlQuery;


        }
    }
}
