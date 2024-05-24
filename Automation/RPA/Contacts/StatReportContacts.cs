using ClosedXML.Excel;
using SAP.Middleware.Connector;
using System;
using System.Data;
using System.IO;
using System.Linq;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using System.Collections.Generic;
using DataBotV5.Data.Stats;

namespace DataBotV5.Automation.RPA.Contacts
{    /// <summary><c>ContactsCreation:</c> 
     /// Clase RPA Automation encargada de enviar reporte de estadísticas.</summary>
    class StatReportContacts
    {
        Rooting root = new Rooting();
        ConsoleFormat console = new ConsoleFormat();
        MailInteraction mail = new MailInteraction();
        Credentials cred = new Credentials();
        ValidateData val = new ValidateData();

        Database DB = new Database();
        SapVariants sap = new SapVariants();
        CRUD crud = new CRUD();

        string respFinal = "";

        public void Main()
        {
            string mand = "QAS";
            string mandante = "CRM";
            string erpSystem = "ERP";

            //nuevas solicitudes de reporte
            DataTable contactReport = null;
            #region Leer solicitudes
            try
            {
                string sql3 = "SELECT DISTINCT `userName`,`userRol` FROM `ReportHistory` WHERE status = 1 ORDER BY ReportHistory.`userName` ASC";
                contactReport = crud.Select( sql3, "update_contacts");
            }
            catch (Exception ex)
            {
                //auto apagar
                try
                {
                    string sqlU = $"UPDATE `orchestrator` SET `active`= 0 WHERE `class` = '{root.BDMethod}'";
                    crud.Update(sqlU, "databot_db");
                }
                catch (Exception ex2)
                { }
                mail.SendHTMLMail("No se pudo conectar con la BD de S&S<br><br>" + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, "Error en crear reporte de estadiscticas de S&S<br>", new string[] { "dmeza@gbm.net", "joarojas@gbm.net" });
            }
            #endregion

            try
            {
                if (contactReport.Rows.Count > 0)
                {
                    console.WriteLine("Procesando...");

                    //confirm contact info
                    DataTable confirmContact = new DataTable();
                    //empleados databot
                    DataTable empleados = new DataTable();
                    //empleados name
                    DataTable emploName = new DataTable();
                    console.WriteLine("Get Employees...");
                    emploName = crud.Select("SELECT * FROM `digital_sign` ORDER BY `user` DESC ", "MIS");
                    //clientes databot
                    DataTable customers = new DataTable();

                    //crea el dataTable customersReport que será el excel 
                    DataTable customersReport = new DataTable();
                    customersReport.Columns.Add("Cliente");
                    customersReport.Columns.Add("Nombre");
                    customersReport.Columns.Add("País");
                    customersReport.Columns.Add("Territorio");
                    customersReport.Columns.Add("Empleado Responsable");
                    customersReport.Columns.Add("Gerente Responsable");
                    customersReport.Columns.Add("Cantidad de Contactos");
                    customersReport.Columns.Add("Cantidad de Contactos Confirmados");
                    customersReport.Columns.Add("Porcentaje de Cumplimiento", typeof(double));

                    //crea el datatable statsReport que será el excel
                    DataTable statsReport = new DataTable();
                    statsReport.Columns.Add("Colaborador");
                    statsReport.Columns.Add("Usuario");
                    statsReport.Columns.Add("País");
                    statsReport.Columns.Add("Total de Clientes");
                    statsReport.Columns.Add("Total de Contactos");
                    statsReport.Columns.Add("Total de contactos confirmados");
                    statsReport.Columns.Add("Porcentaje de Cumplimiento", typeof(double));

                    //crea el datatable statsReport que será el excel
                    DataTable statsSalesRepReport = new DataTable();
                    statsSalesRepReport.Columns.Add("País");
                    statsSalesRepReport.Columns.Add("Total de Clientes");
                    statsSalesRepReport.Columns.Add("Total de Contactos");
                    statsSalesRepReport.Columns.Add("Total de Contactos confirmados");
                    statsSalesRepReport.Columns.Add("Porcentaje de Cumplimiento", typeof(double));

                    //crea el datatable historyReport que será el excel y le agrega las columnas de la tabla HistoryContacts
                    DataTable historyContactColumns = crud.Select( "show full COLUMNS from `HistoryContacts`", "update_contacts");
                    DataTable historyReport = new DataTable();
                    historyContactColumns.Rows.Cast<DataRow>().ToList().ForEach(history =>
                    {
                        historyReport.Columns.Add(history["Field"].ToString());
                    });
                    historyReport.Columns["Create_Date"].DataType = System.Type.GetType("System.DateTime");
                    historyReport.Columns["Update_Date"].DataType = System.Type.GetType("System.DateTime");


                    string filas, res1, res2, mensaje;
                    filas = res1 = res2 = mensaje = "";

                    foreach (DataRow solicitud in contactReport.Rows)
                    {
                        XLWorkbook wb = new XLWorkbook();
                        //DataRow solicitud = contactReport.Select("userName = '" + user["userName"] + "'")[0];

                        string rol = solicitud["userRol"].ToString();
                        string userName = solicitud["userName"].ToString();
                        string userId = "";
                        string userCountry = "";

                        #region Crear reporte
                        console.WriteLine("Set History Report...");
                        string historyQuery = "";
                        string confirmQuery = "";
                        string EmpQuery = "";
                        if (rol == "V")
                        {
                            customersReport.Clear();
                            historyQuery = $"SELECT * FROM `HistoryContacts` WHERE Create_By = '{userName}'";
                            confirmQuery = $"SELECT * FROM `ConfirmContacts` WHERE CreateBy = '{userName}'";
                            EmpQuery = $"SELECT DISTINCT `accountManagerUser`, `accountManagerId`, `manager` FROM `clients` WHERE accountManagerUser = '{userName}' ORDER BY accountManagerUser ASC";
                            //por lo que se busca en empleados
                            empleados = crud.Select($"SELECT * FROM `digital_sign` WHERE user = '{userName}'", "MIS");
                       
                            userId = empleados.Rows[0]["UserID"].ToString();
                            userCountry = empleados.Rows[0]["name"].ToString();

                            statsSalesRepReport.Clear();
                            //crear reporte de history change
                            DataTable historySalesRepReport = new DataTable();
                            historyContactColumns.Rows.Cast<DataRow>().ToList().ForEach(history =>
                            {
                                historySalesRepReport.Columns.Add(history["Field"].ToString());
                            });

                            historySalesRepReport.Columns["Create_Date"].DataType = System.Type.GetType("System.DateTime");
                            historySalesRepReport.Columns["Update_Date"].DataType = System.Type.GetType("System.DateTime");

                            historySalesRepReport = crud.Select( historyQuery, "update_contacts");

                            //despues de agregar el log se cambia las columnas para que salga bien en el excel
                            //por lo que se toma el comentario de la columna del historyContacts para agregarlo como titulo de la columna del Excel
                            historyContactColumns.Rows.Cast<DataRow>().ToList().ForEach(history =>
                            {
                                historySalesRepReport.Columns[history["Field"].ToString()].ColumnName = history["Comment"].ToString();
                            });

                            wb.Worksheets.Add(historySalesRepReport, "Historial Cambios");

                            //Extrae los contactos confirmados
                            confirmContact = crud.Select( confirmQuery, "update_contacts");

                        }
                        else
                        {
                            //gerentes y admis
                            historyQuery = "SELECT * FROM `HistoryContacts`";
                            confirmQuery = "SELECT * FROM `ConfirmContacts`";
                            EmpQuery = "SELECT DISTINCT `accountManagerUser`, `accountManagerId`, `manager` FROM `clients` ORDER BY accountManagerUser ASC";
                            empleados = crud.Select(EmpQuery, "databot_db");

                            //crear reporte de history change
                            if (historyReport.Rows.Count <= 0)
                            {
                                historyReport = crud.Select( historyQuery, "update_contacts");

                                //despues de agregar el log se cambia las columnas para que salga bien en el excel
                                //por lo que se toma el comentario de la columna del historyContacts para agregarlo como titulo de la columna del Excel
                                historyContactColumns.Rows.Cast<DataRow>().ToList().ForEach(history =>
                                {
                                    historyReport.Columns[history["Field"].ToString()].ColumnName = history["Comment"].ToString();
                                });
                            }

                            if (confirmContact.Rows.Count <= 0)
                            {
                                //Extrae los contactos confirmados
                                confirmContact = crud.Select( confirmQuery, "update_contacts");
                            }
                        }

                        //crear reporte de estadisticas
                        // crea el reporte si el rol es vendedor o bien si es admin y nunca se ha creado el reporte en un usuario previo
                        if (rol == "V" || rol == "A" && statsReport.Rows.Count <= 0)
                        {
                            console.WriteLine("Set Stats Report...");
                            empleados.Rows.Cast<DataRow>().ToList().ForEach(empleado =>
                            {
                                string salesRepId = empleado["UserID"].ToString().PadLeft(8, '0');

                                salesRepId = (salesRepId.Substring(0, 2) != "AA") ? "AA" + salesRepId : salesRepId;

                                string salesRepUser = empleado["user"].ToString();
                                string salesRepName = "", salesRepCountry = "";
                                try
                                {
                                    salesRepName = empleado["name"].ToString();
                                    salesRepCountry = empleado["country"].ToString();
                                }
                                catch (Exception)
                                {
                                    try
                                    {
                                        salesRepName = emploName.Select($"user = '{salesRepUser}'")[0]["name"].ToString();
                                        salesRepCountry = emploName.Select($"user = '{salesRepUser}'")[0]["country"].ToString();
                                    }
                                    catch {; }
                                }
                                //por cada empleado correr la cantidad de cliente y contactos asignados
                                int totalCustomers = 0;
                                double totalContacts = 0;
                                try
                                {
                                    Dictionary<string, string> parameters = new Dictionary<string, string>();
                                    parameters["IDSALESREP"] = salesRepId;

                                    IRfcFunction getTotals = sap.ExecuteRFC(mandante, "ZDM_AM_GET_TOTALS", parameters);


                                    totalCustomers = (int)getTotals.GetValue("TOTAL_CLIENTES");
                                    totalContacts = (int)getTotals.GetValue("TOTAL_CONTACTOS");
                                }
                                catch (Exception)
                                {
                                    string err = "";
                                }
                                if (totalCustomers == 0)
                                {
                                    return;
                                }
                                //sacar cuantos contactos se han confirmado
                                DataRow[] contactsConfirm = new DataRow[0];
                                try { contactsConfirm = confirmContact.Select($"CreateBy = '{salesRepUser}'"); } catch {; }
                                double cantConfirm = contactsConfirm.Length;
                                double porcent = 0;
                                if (totalContacts != 0)
                                {
                                    //sacar el % de cumplimiento
                                    porcent = ((cantConfirm * 100) / totalContacts);
                                }
                                if (porcent != 0)
                                {
                                    porcent = Math.Round(porcent, 2);
                                }

                                //llenar excel:
                                //add row to excel
                                if (rol == "V")
                                {

                                    DataRow statRow = statsSalesRepReport.Rows.Add();
                                    statRow["País"] = salesRepCountry.ToString();
                                    statRow["Total de Clientes"] = totalCustomers.ToString();
                                    statRow["Total de Contactos"] = totalContacts.ToString();
                                    statRow["Total de contactos confirmados"] = cantConfirm.ToString();
                                    statRow["Porcentaje de Cumplimiento"] = porcent;
                                    statsSalesRepReport.AcceptChanges();
                                }
                                else
                                {
                                    DataRow statRow = statsReport.Rows.Add();
                                    statRow["Colaborador"] = salesRepName;
                                    statRow["Usuario"] = salesRepUser;
                                    statRow["País"] = salesRepCountry;
                                    statRow["Total de Clientes"] = totalCustomers.ToString();
                                    statRow["Total de Contactos"] = totalContacts.ToString();
                                    statRow["Total de contactos confirmados"] = cantConfirm.ToString();
                                    statRow["Porcentaje de Cumplimiento"] = porcent;
                                    statsReport.AcceptChanges();
                                }




                            });

                        }

                        //Save sheets

                        try { historyReport.Columns.Remove("Columnas modificadas"); } catch (Exception) { }

                        if (rol == "V")
                        {
                            DataView dv = statsSalesRepReport.DefaultView;
                            dv.Sort = "Porcentaje de Cumplimiento asc";
                            statsSalesRepReport = dv.ToTable();
                            wb.Worksheets.Add(statsSalesRepReport, "Estadisticas");
                            //el historial se agrega arriba
                        }
                        else
                        {
                            DataView dv = statsReport.DefaultView;
                            dv.Sort = "Porcentaje de Cumplimiento DESC";
                            statsReport = dv.ToTable();
                            wb.Worksheets.Add(statsReport, "Estadisticas");

                            wb.Worksheets.Add(historyReport, "Historial Cambios");

                        }

                        //crear reporte de customers y su % de cumplimiento

                        if (rol == "V")
                        {

                            Dictionary<string, string> parameters = new Dictionary<string, string>();
                            parameters["IDSALESREP"] = userId;

                            IRfcFunction getCustomers = sap.ExecuteRFC(erpSystem, "ZDM_CUSTOMERS_SALESREP", parameters);

                            customers = sap.GetDataTableFromRFCTable(getCustomers.GetTable("CUSTOMERS"));
                        }
                        else
                        {
                            customers = crud.Select(@"SELECT clients.*, sapCountries.countryCode, valueTeam.valueTeam
                                                        FROM `clients` 
                                                        INNER JOIN sapCountries ON sapCountries.id = clients.country
                                                        INNER JOIN valueTeam ON valueTeam.id = clients.territory
                                                        ORDER BY accountManagerUser ASC", "databot_db");

                            customers.Columns["accountManagerUser"].ColumnName = "ID_ERP_CUSTOMER";
                            customers.Columns["name"].ColumnName = "NAME";
                            customers.Columns["countryCode"].ColumnName = "COUNTRY";
                            customers.Columns["valueTeam"].ColumnName = "TERRITORY";
                            customers.Columns["accountManagerUser"].ColumnName = "SALES_REP";
                            customers.Columns["manager"].ColumnName = "MANAGER";
                            customers.Columns["locked"].ColumnName = "INACTIVE";

                        }

                        //el reporte de admin tiene muchos clientes por lo que se deja solo para vendedores
                        if (rol == "V")
                        {

                            console.WriteLine("Set Customer Report...");

                            customers.Rows.Cast<DataRow>().ToList().ForEach(customer =>
                            {
                                if (customer["INACTIVE"].ToString() != "X")
                                {
                                    string customerId = customer["ID_ERP_CUSTOMER"].ToString().PadLeft(10, '0');
                                    int totalContact = 0;
                                    try
                                    {
                                        Dictionary<string, string> parameters = new Dictionary<string, string>();
                                        parameters["IDSALESREP"] = customerId;

                                        IRfcFunction getTotals = sap.ExecuteRFC(mandante, "ZDM_AM_GET_TOTALS", parameters);

                                        totalContact = (int)getTotals.GetValue("TOTAL_CONTACTOS");
                                    }
                                    catch (Exception)
                                    { }

                                    //int totalCustomers = (int)getTotals.GetValue("TOTAL_CLIENTES");

                                    //sacar cuantos contactos se han confirmado
                                    DataRow[] contactConfirm = new DataRow[0];
                                    try { contactConfirm = confirmContact.Select($"idCustomer = '{customerId}'"); } catch {; }
                                    double cantConfirmed = contactConfirm.Length;
                                    double percent = 0;
                                    if (totalContact != 0)
                                    {
                                        //sacar el % de cumplimiento
                                        percent = ((cantConfirmed * 100) / totalContact);
                                    }
                                    if (percent != 0)
                                    {
                                        percent = Math.Round(percent, 2);
                                    }

                                    DataRow custRow = customersReport.Rows.Add();
                                    custRow["Cliente"] = customerId;
                                    custRow["Nombre"] = customer["NAME"].ToString();
                                    custRow["País"] = customer["COUNTRY"].ToString();
                                    custRow["Territorio"] = customer["TERRITORY"].ToString();
                                    custRow["Empleado Responsable"] = customer["SALES_REP"].ToString();
                                    custRow["Gerente Responsable"] = customer["MANAGER"].ToString();
                                    custRow["Cantidad de Contactos"] = totalContact;
                                    custRow["Cantidad de Contactos Confirmados"] = cantConfirmed;
                                    custRow["Porcentaje de Cumplimiento"] = percent;
                                    customersReport.AcceptChanges();
                                }
                            });

                            DataView dv2 = customersReport.DefaultView;
                            dv2.Sort = "Porcentaje de Cumplimiento DESC";
                            customersReport = dv2.ToTable();
                            wb.Worksheets.Add(customersReport, "Avance por Cliente");
                        }

                        #endregion

                        #region save excel
                        console.WriteLine("Save Excel...");

                        string fecha_file = $"{DateTime.Now.Day.ToString().PadLeft(2, '0')}_{DateTime.Now.Month.ToString().PadLeft(2, '0')}_{DateTime.Now.Year}";

                        string ruta = (rol == "V") ? root.FilesDownloadPath + $"\\Reporte Estadisticas Contactos {fecha_file} {userName}.xlsx" : root.FilesDownloadPath + $"\\Reporte Estadisticas Contactos {fecha_file}.xlsx";

                        if (File.Exists(ruta))
                        {
                            File.Delete(ruta);
                        }
                        wb.SaveAs(ruta);

                        //to send email
                        console.WriteLine("Send Email...");
                        string[] cc = { "" };
                        string[] adj = { ruta };
                        string sub = "Reporte Estadisticas - Actualización de contactos - " + fecha_file.Replace("_", "/");
                        string msj = "A continuación, se adjunta las estadisticas y log de cambios en el portal de Actualización de Contactos a la fecha " + DateTime.Today.ToString();
                        string html = Properties.Resources.emailtemplate1;
                        html = html.Replace("{subject}", "Reporte de estadísticas - actualización de contactos");
                        html = html.Replace("{cuerpo}", msj);
                        html = html.Replace("{contenido}", "");
                        #endregion

                        #region Enviar notificacion al solicitante

                        //contactReport.Rows.Cast<DataRow>().ToList().ForEach(dataRow =>
                        //{
                        //cambiar status
                        string sqlU = $"UPDATE `ReportHistory` SET `status` = '0', `updateAt` = '{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}' WHERE `userName` = '{userName}'";
                        crud.Update(sqlU, "update_contacts");

                        //enviar reporte a cada persona dentro de las solicitudes
                        mail.SendHTMLMail(html, new string[] { userName.ToString() + "@gbm.net" }, sub, cc, adj);

                        root.BDUserCreatedBy = userName.ToString();
                        using (Stats stats = new Stats())
                        {
                            stats.CreateStat();
                        }
                        //});
                        #endregion
                    }



                }

            }
            catch (Exception exs)
            {
                //auto apagar
                try
                {
                    crud.Update($"UPDATE `orchestrator` SET `active`= 0 WHERE `class` = '{root.BDMethod}'", "databot_db");
                }
                catch (Exception ex2)
                { }

                mail.SendHTMLMail("Error al crear reporte:<br><br>" + exs, new string[] { "internalcustomersrvs@gbm.net" }, "Error en crear reporte de estadiscticas de S&S<br>", new string[] { "dmeza@gbm.net", "joarojas@gbm.net" });

            }
        }
    }

}
