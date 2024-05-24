using DataBotV5.Data.Projects.MasterData;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Data.Database;
using DataBotV5.Logical.Webex;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Process;

using DataBotV5.Logical.Web;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using System.Globalization;
using Newtonsoft.Json.Linq;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Threading;
using ClosedXML.Excel;
using System.Data;
using System;
using DataBotV5.Logical.MicrosoftTools;

namespace DataBotV5.Automation.DM.Warranties
{
    /// <summary>
    /// Clase DM Automation encargada de enviar garantías de datos maestros vía correo electrónico.
    /// </summary>
    /// 


    class NewWarrantyRequestSS
    {
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        MasterDataSqlSS DM = new MasterDataSqlSS();
        SapVariants sap = new SapVariants();
        Rooting root = new Rooting();
        MsExcel ms = new MsExcel();
        Stats stats = new Stats();
        CRUD crud = new CRUD();

        Log log = new Log();

        string crmMand = "CRM";
        //int mandCrm = 460;

        string respFinal = "";


        public void Main()
        {
            string respuesta = DM.GetManagement("10"); //GARANTIAS
            if (!String.IsNullOrEmpty(respuesta) && respuesta != "ERROR")
            {
                console.WriteLine("Procesando...");
                ProcessWarranty();

                console.WriteLine("Creando Estadísticas");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }

        }
        public void ProcessWarranty()
        {
            string[] cc = { "hlherrera@gbm.net" };

            try
            {
                root.requestDetails = root.requestDetails.Replace("\u00A0", " "); //eliminar non breaks spaces (char 160)
                root.requestDetails = root.requestDetails.Replace(@"[^\u0000-\u007F]+", ""); //eliminar caracteres no ASCII

                if (root.metodoDM != "1") //lineal
                {
                    //no hay lineal, por el momento
                    mail.SendHTMLMail("Gestion_garantia: " + root.IdGestionDM + "<br>Solicitante_garantia: " + root.BDUserCreatedBy + @"<br>ARCHIVO: adjunto" + root.IdGestionDM, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject, cc, new string[] { root.FilesDownloadPath + "\\" + root.ExcelFile });
                    DM.ChangeStateDM(root.IdGestionDM, "", "14"); //PENDIENTE
                }
                else //MASIVO
                {
                    //el metodo masivo todavia no funciona, se le envia un correo a ICS con la plantilla
                    #region abrir excel

                    Console.WriteLine(DateTime.Now + " > > > " + "Abriendo excel y validando");
                    string adjunto = root.ExcelFile;
                    DataTable xlWorkSheet = ms.GetExcel(root.FilesDownloadPath + "\\" + adjunto);
                    int rows = xlWorkSheet.Rows.Count;
                    #endregion

                    DataTable datos = new DataTable();
                    datos.Columns.Add("EQUNR");
                    datos.Columns.Add("MATNR");
                    datos.Columns.Add("SERNR");


                    List<string> idsList = new List<string>();
                    List<string> seriesList = new List<string>();
                    List<string> materialesList = new List<string>();

                    foreach (DataRow row in xlWorkSheet.Rows)
                    {
                        idsList.Add(row["ID OBJETO"].ToString());
                        seriesList.Add(row["SERIE"].ToString());
                        materialesList.Add(row["MATERIAL"].ToString());
                    }


                    string[] ids = idsList.ToArray();
                    string[] series = seriesList.ToArray();
                    string[] materiales = materialesList.ToArray();

                    if (series.Length != materiales.Length)
                    {
                        // error faltan materiales o series?
                    }

                    //RfcDestination destination = SAP.Middleware.Connector.RfcDestinationManager.GetDestination(cred.parametros);
                    //RfcRepository repo = destination.Repository;

                    RfcDestination destCrm = sap.GetDestRFC(crmMand);
                    // IRfcFunction fmMg = destCrm.Repository.CreateFunction("RFC_READ_TABLE");

                    #region traer datos de Equi



                    //IRfcFunction zdmCreateMat = sap.ExecuteRFC(crmMand, "ZDM_CREATE_MAT", zdmCreateMatParameters);// ver esto


                    IRfcFunction fmTable = destCrm.Repository.CreateFunction("RFC_READ_TABLE");
                    fmTable.SetValue("USE_ET_DATA_4_RETURN", "X");
                    fmTable.SetValue("QUERY_TABLE", "EQUI");
                    fmTable.SetValue("DELIMITER", "|");
                    //fmTable.SetValue("GET_SORTED", "X");
                    //fmTable.SetValue("ROWCOUNT", "100");

                    IRfcTable fields = fmTable.GetTable("FIELDS");

                    fields.Append();
                    fields.SetValue("FIELDNAME", "SERNR");           //campos a traer
                    fields.Append();
                    fields.SetValue("FIELDNAME", "MATNR");
                    fields.Append();
                    fields.SetValue("FIELDNAME", "EQUNR");


                    IRfcTable fm_optionsEqui = fmTable.GetTable("OPTIONS");

                    for (int i = 0; i <= series.Length - 1; i++)
                    {
                        if (series[i] != "" && materiales[i] != "" && series[i] != "PB018ZZH")             // a veces el excel trae líneas en blanco y afecta el query. Además de obviar la línea de ejemplo
                        {
                            fm_optionsEqui.Append();
                            fm_optionsEqui.SetValue("TEXT", "( SERNR IN ('" + series[i] + "' ) AND MATNR IN ('" + materiales[i] + "' ) )");  //tener cuidado de no sobrepasar los 72 caracteres de longitud que acepta la FM por línea

                            if (i != series.Length - 1)
                            {
                                fm_optionsEqui.Append();
                                fm_optionsEqui.SetValue("TEXT", " OR ");        // hacer consulta de siguiente línea de la solicitud
                            }
                        }
                    }

                    fmTable.Invoke(destCrm);

                    IRfcFunction fmTableMara = destCrm.Repository.CreateFunction("RFC_READ_TABLE");
                    fmTableMara.SetValue("USE_ET_DATA_4_RETURN", "X");
                    fmTableMara.SetValue("QUERY_TABLE", "MARA");
                    fmTableMara.SetValue("DELIMITER", "|");
                    //fmTable.SetValue("GET_SORTED", "X");
                    //fmTable.SetValue("ROWCOUNT", "100");

                    IRfcTable fieldsMara = fmTableMara.GetTable("FIELDS");

                    fieldsMara.Append();
                    fieldsMara.SetValue("FIELDNAME", "MATNR");
                    fieldsMara.Append();
                    fieldsMara.SetValue("FIELDNAME", "MATKL");           //campos a traer


                    IRfcTable fm_optionsMara = fmTableMara.GetTable("OPTIONS");

                    for (int i = 0; i <= materiales.Length - 1; i++)
                    {
                        if (materiales[i] != "" && series[i] != "PB018ZZH")             // a veces el excel trae líneas en blanco y afecta el query. Además de obviar la línea de ejemplo
                        {
                            fm_optionsMara.Append();
                            fm_optionsMara.SetValue("TEXT", "MATNR IN ('" + materiales[i] + "' )");  //tener cuidado de no sobrepasar los 72 caracteres de longitud que acepta la FM por línea

                            if (i != materiales.Length - 1)
                            {
                                fm_optionsMara.Append();
                                fm_optionsMara.SetValue("TEXT", " OR ");        // hacer consulta de siguiente línea de la solicitud
                            }

                        }
                    }

                    fmTableMara.Invoke(destCrm);



                    DataTable reporteMara = GetDataTableFromRFCTable(fmTableMara.GetTable("ET_DATA"));

                    DataTable reporte = GetDataTableFromRFCTable(fmTable.GetTable("ET_DATA"));

                    foreach (DataRow fila in GetDataTableFromRFCTable(fmTable.GetTable("ET_DATA")).Rows)
                    {
                        DataRow filaReporte = datos.NewRow();
                        filaReporte["EQUNR"] = fila["LINE"].ToString().Split(new char[] { '|' })[0].Trim().TrimStart(new char[] { '0' });
                        filaReporte["MATNR"] = fila["LINE"].ToString().Split(new char[] { '|' })[1].Trim();
                        filaReporte["SERNR"] = fila["LINE"].ToString().Split(new char[] { '|' })[2].Trim();

                        datos.Rows.Add(filaReporte);
                    }

                    datos.Columns.Add("CISCO");

                    foreach (DataRow fila in reporteMara.Rows)
                    {
                        foreach (DataRow rowDatos in datos.Rows)
                        {
                            if (rowDatos.ItemArray[1].ToString() == fila["LINE"].ToString().Split(new char[] { '|' })[0].Trim() && fila["LINE"].ToString().Split(new char[] { '|' })[1].Trim().Substring(0, 3) == "103")
                            {
                                rowDatos["CISCO"] = "X";
                            }
                        }
                    }

                    #endregion



                    IRfcFunction Z_TEST_GET_WARRANTIES = destCrm.Repository.CreateFunction("Z_TEST_GET_WARRANTIES");
                    IRfcTable id_equipment = Z_TEST_GET_WARRANTIES.GetTable("ID_EQUIPMENT");
                    IRfcTable resultadoFM;


                    foreach (DataRow rowEqui in datos.Rows)
                    {
                        id_equipment.Append();                                       //se agrega una línea a la tabla EQ_ID
                        id_equipment.SetValue("EQ_ID", rowEqui.ItemArray[0]);        //AQUÍ SE LLENÓ UNA LÍNEA y se repite si se necesita más
                    }

                    Z_TEST_GET_WARRANTIES.Invoke(destCrm);

                    resultadoFM = Z_TEST_GET_WARRANTIES.GetTable("RESULTADO");

                    DataTable resultadoFMWarranties = GetDataTableFromRFCTable(resultadoFM);

                    resultadoFMWarranties.Columns.Add("CISCO");
                    resultadoFMWarranties.Columns.Remove("ITEM_PROD");
                    resultadoFMWarranties.Columns.Remove("ITEM_NO");


                    foreach (DataRow rowResultado in resultadoFMWarranties.Rows)
                    {
                        foreach (DataRow rowDatos in datos.Rows)
                        {
                            if (rowDatos.ItemArray[0].ToString() == rowResultado.ItemArray[0].ToString() && rowResultado.ItemArray[2].ToString() == "")
                            {
                                rowResultado["EQ_MAT"] = rowDatos.ItemArray[1];
                                rowResultado["EQ_SER"] = rowDatos.ItemArray[2];

                                if (rowDatos.ItemArray[3].ToString().Trim() == "X")
                                {
                                    rowResultado["CISCO"] = "X";
                                }

                                break;
                            }
                        }
                    }

                    //resultadoFMWarranties = resort(resultadoFMWarranties, "EQ_ID", "DESC");

                    CultureInfo culture = (CultureInfo)CultureInfo.CurrentCulture.Clone();
                    culture.DateTimeFormat.ShortDatePattern = "dd.MM.yyyy";
                    culture.DateTimeFormat.LongTimePattern = "";
                    Thread.CurrentThread.CurrentCulture = culture;

                    bool fixDate = false, notificar = false, validFabricWarranty = false, validDate = false;
                    string fabWarranty, warrantyRequested; DateTime startDateRequested; DateTime currentStartDate, endDateRequested, currentEndDate; double differenceDate;

                    foreach (DataRow row in xlWorkSheet.Rows)
                    {
                        int i = xlWorkSheet.Rows.IndexOf(row);
                        startDateRequested = Convert.ToDateTime(row["INICIO GARANTIA"].ToString());
                        endDateRequested = Convert.ToDateTime(row["FIN GARANTIA"].ToString());
                        warrantyRequested = row["TIPO GARANTIA"].ToString().Substring(11);


                        //DataTable warrantyRange = crud.Select("Databot", "select `WARRANTY_RANGE` from warranties_descriptions where DESCRIPTION = ''" + warrantyRequested + "'", "warranties", "DEV");
                        //string warrantyRange1 = crud.Select("Databot", "select `WARRANTY_RANGE` from warranties_descriptions where DESCRIPTION = ''" + warrantyRequested + "'", "warranties", "DEV").Rows[0].ItemArray[0].ToString();
                        //string warrantyRange1 = db.SelectRow("automation", "SELECT `WARRANTY_RANGE` FROM `warranties`" +
                        //                            " WHERE DESCRIPTION = '" + warrantyRequested + "'").Rows[0].ItemArray[0].ToString();

                        if (resultadoFMWarranties.Rows[i]["WARR_START"].ToString() != "00000000")
                        {
                            currentStartDate = DateTime.ParseExact(resultadoFMWarranties.Rows[i]["WARR_START"].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                            currentEndDate = DateTime.ParseExact(resultadoFMWarranties.Rows[i]["WARR_END"].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);

                        }
                        else if (resultadoFMWarranties.Rows[i]["CISCO"].ToString() == "X")
                        {
                            if (warrantyRequested.Contains("EXT"))
                            {
                                //if (Convert.ToDouble(warrantyRange) <= Math.Round((endDateRequested - startDateRequested).TotalDays) / 365)
                                //{
                                //    //agregar garantía ext cisco

                                //}
                            }
                            else
                            {
                                //la garantía no es Ext
                            }

                        }
                        else if (!warrantyRequested.Contains("EXT"))
                        {
                            //if (Convert.ToDouble(warrantyRange) <= Math.Round((endDateRequested - startDateRequested).TotalDays) / 365)
                            //{
                            //    //agregar garantía fábrica

                            //}
                        }
                        else if (resultadoFMWarranties.Rows[i]["CISCO"].ToString() != "X")
                        {
                            // el equipo no tiene garantía de fábrica, rechazar.
                        }

                        differenceDate = Math.Round(((endDateRequested - startDateRequested).TotalDays) / 365);

                        //cred.IngresarAmbiente(mandante);

                        mail.SendHTMLMail("Gestion_garantia: " + root.IdGestionDM + "<br>Solicitante_garantia: " + root.BDUserCreatedBy + @"<br>ARCHIVO: \\RPAWEB\Users\Administrator\Desktop\FTP_FILES\dm_gestiones_mass\" + root.IdGestionDM, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject, cc/*, adjunto*/);
                        DM.ChangeStateDM(root.IdGestionDM, "", "14"); //PENDIENTE

                        log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Gestión de garantía", root.IdGestionDM, root.Subject);
                        respFinal = respFinal + "\\n" + "Gestión garantía: " + root.IdGestionDM + " Solicitante garantía: " + root.BDUserCreatedBy;




                    }



                }

                root.requestDetails = respFinal;

            }
            catch (Exception ex)
            {
                DM.ChangeStateDM(root.IdGestionDM, ex.Message, "4"); //ERROR
                mail.SendHTMLMail("Gestion: " + root.IdGestionDM + "<br>" + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, cc);
            }
        }

        public static DataTable GetDataTableFromRFCTable(IRfcTable lrfcTable)
        {
            //sapnco_util
            DataTable loTable = new DataTable();

            //... Create ADO.Net table.
            for (int liElement = 0; liElement < lrfcTable.ElementCount; liElement++)
            {
                RfcElementMetadata metadata = lrfcTable.GetElementMetadata(liElement);
                if (metadata.DataType.ToString() == "TABLE")
                    loTable.Columns.Add(metadata.Name, typeof(DataTable));
                else
                    loTable.Columns.Add(metadata.Name);
            }

            //... Transfer rows from lrfcTable to ADO.Net table.
            foreach (IRfcStructure row in lrfcTable)
            {
                DataRow ldr = loTable.NewRow();
                for (int liElement = 0; liElement < lrfcTable.ElementCount; liElement++)
                {
                    RfcElementMetadata metadata = lrfcTable.GetElementMetadata(liElement);
                    try { ldr[metadata.Name] = row.GetString(metadata.Name); }
                    catch (Exception)
                    {
                        //ldr[metadata.Name] = "Es otra Tabla, por favor tomarla aparte";
                    }
                }
                loTable.Rows.Add(ldr);
            }
            return loTable;
        }

    }
}


