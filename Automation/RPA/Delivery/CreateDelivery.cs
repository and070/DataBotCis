using System;
using Excel = Microsoft.Office.Interop.Excel;
using SAP.Middleware.Connector;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using DataBotV5.Data.SAP;
using System.Collections.Generic;
using DataBotV5.Logical.MicrosoftTools;
using System.Data;
using System.Linq;
using DataBotV5.Data.Database;
using DataBotV5.App.ConsoleApp;
using System.Collections;

namespace DataBotV5.Automation.RPA.Delivery
{
    /// <summary>
    /// Clase RPA Automation encargada de la creación de delivery.
    /// </summary>
    class CreateDelivery
    {
        #region variables globales
        ConsoleFormat console = new ConsoleFormat();
        public string response = "";
        public string response_failure = "";
        Credentials cred = new Credentials();
        MailInteraction mail = new MailInteraction();
        CRUD crud = new CRUD();
        Rooting root = new Rooting();
        ValidateData val = new ValidateData();
        ProcessInteraction proc = new ProcessInteraction();
        Log log = new Log();
        Stats estadisticas = new Stats();
        string systemSap = "ERP";
        int mandante = 260;
        SapVariants sap = new SapVariants();
        MsExcel xs = new MsExcel();
        Settings sett = new Settings();
        string enviroment = Start.enviroment;
        //string enviroment = "QAS";
        bool executeStats = false;

        #endregion
        public void Main()
        {

            console.WriteLine("Procesando...");

            ProcessDelivery();
            response = "";

            if (executeStats)
            {
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }

        }
        public void ProcessDelivery()
        {

            #region Variables Privadas
            int rows;
            bool validar_lineas = true;
            bool fmError = false;

            DataTable wb = new DataTable();
            #endregion

            #region Extracción de material requests de hoy a crear deliveries.

            string currentDate = DateTime.Now.Date.ToString("yyyy-MM-dd");
            string monthAgoDate = DateTime.Now.Date.AddMonths(-1).ToString("yyyy-MM-dd");
            string tomorrowDate = DateTime.Now.Date.AddDays(1).ToString("yyyy-MM-dd");


            RfcDestination dest_erp = sap.GetDestRFC(systemSap, mandante);
            IRfcFunction getMr = dest_erp.Repository.CreateFunction("ZFI_GET_MR");

            #region Definir filtros.
            DataRow filters = crud.Select("select * from FMFilters where active=1", "delivery_db", enviroment).Rows[0]; 

            string fromDateParsed = "";
            string toDateParsed = "";

            try
            {
                DateTime.Parse(filters["fromDate"].ToString()).ToString("yyyy-MM-dd");
                DateTime.Parse(filters["toDate"].ToString()).ToString("yyyy-MM-dd");
            }
            catch (Exception e) { }


            string fromDateF = filters["fromDate"].ToString() == "" ? monthAgoDate : fromDateParsed;
            string toDateF = filters["toDate"].ToString() == "" ? tomorrowDate : toDateParsed;
            string plantF = filters["plant"].ToString();
            string plantEndsWithF =  filters["plantEndsWith"].ToString();
            string storageLocationF = filters["storageLocation"].ToString();
            string storageLocationEndsWithF = filters["storageLocationEndsWith"].ToString();
            string excludeDeliveriesF = filters["excludeDeliveries"].ToString();
            string excludeEmptiesMrF = filters["excludeEmptiesMr"].ToString();
            string excludeEmptiesOsF = filters["excludeEmptiesOs"].ToString();
            string deliveryBlockF = filters["deliveryBlock"].ToString();
            string itemBlockF = filters["itemBlock"].ToString();
            #endregion

            getMr.SetValue("FROM_DATE", fromDateF);
            getMr.SetValue("TO_DATE", toDateF);
            getMr.SetValue("PLANT", plantF);
            getMr.SetValue("PLANT_ENDS_WITH", plantEndsWithF);
            getMr.SetValue("STORAGE_LOCATION", storageLocationF);
            getMr.SetValue("STORAGE_LOCATION_ENDS_WITH", storageLocationEndsWithF);
            getMr.SetValue("EXCLUDE_DELIVERIES", excludeDeliveriesF); //Los que ya tienen delivery no lo trae.
            getMr.SetValue("EXCLUDE_EMPTIES_MR", excludeEmptiesMrF); //Los que no tienen mr no lo trae.
            getMr.SetValue("EXCLUDE_EMPTIES_OS", excludeEmptiesOsF); //Los que no tienen os no lo trae.
            getMr.SetValue("DELIVERY_BLOCK", deliveryBlockF); //Flag de cancelación en caso que se quiera bloquear manualmente por tema comercial, otro fin urgente, por decisión comercial
            getMr.SetValue("ITEM_BLOCK", itemBlockF); ////Flag de cancelación de ítem en caso que se quiera bloquear manualmente por tema comercial, otro fin urgente, por decisión comercial
            console.WriteLine($"Extrayendo Mrs desde {fromDateF} a {toDateF}");
            getMr.Invoke(dest_erp);

            if (getMr.GetValue("RESPONSE").ToString() != "OK")
            {
                //manejar el error 
                fmError = true;
            }

            wb = sap.GetDataTableFromRFCTable(getMr.GetTable("MR_RESULT"));

            #region Filtrar filas para solo las Storage Location 02
            try
            {
                //Filtrar sólo las que tienen storage location terminan en 02 que es la que interesa.
                EnumerableRowCollection<DataRow> filteredRows = wb.AsEnumerable()
                .Where(row => row.Field<string>("STORAGE_LOCATION") != null && row.Field<string>("STORAGE_LOCATION").EndsWith("02"));

                if (filteredRows.Any())
                {
                    // Si hay filas después de aplicar el filtro, copiar a DataTable.
                    wb = filteredRows.CopyToDataTable();
                }
                else
                {
                    wb = new DataTable();
                }

            }
            catch (Exception e)
            {
                console.WriteLine(e.ToString());
            }

            #endregion

            //Hallar deliveries creados manualmente.
            string queryValues = getDeliveriesToAddDataBase(wb);

            #region Filtrar filas para eliminar deliveries creados.
            try
            {

                //Filtrar sólo los que no tienen deliveries.
                EnumerableRowCollection<DataRow> filteredRows = wb.AsEnumerable()
                     .Where(row => string.IsNullOrEmpty(row.Field<string>("DELIVERY")) || string.IsNullOrWhiteSpace(row.Field<string>("DELIVERY")));

                if (filteredRows.Any())
                {
                    // Si hay filas después de aplicar el filtro, copiar a DataTable.
                    wb = filteredRows.CopyToDataTable();
                }
                else
                {
                    wb = new DataTable();
                }

            }
            catch (Exception e)
            {
                console.WriteLine(e.ToString());
            }

            #endregion

            int mrCount = wb.Rows.Count;

            #endregion

            #region Crear la estructura del excel result.

            //Crear el dataTable para el excel de respuesta
            DataTable excelResult = new DataTable();
            try
            {
                excelResult.Columns.Add("Purchase Order");
                excelResult.Columns.Add("Material Request");
                excelResult.Columns.Add("Planta (país)");
                excelResult.Columns.Add("Storage Location)");
                excelResult.Columns.Add("Item");
                excelResult.Columns.Add("Material");
                excelResult.Columns.Add("Cantidad");
                excelResult.Columns.Add("Delivery");
                excelResult.Columns.Add("Order Service");
                excelResult.Columns.Add("Customer PO Order");
                excelResult.Columns.Add("Migo");
                excelResult.Columns.Add("Posting Date");

                excelResult.Columns.Add("ID del Documento");
                excelResult.Columns.Add("WMS Order Transfer");
                excelResult.Columns.Add("Resultado");
            }
            catch (Exception)
            { }

            rows = wb.Rows.Count;


            #endregion

            string respFinal = "";
            string resultPath = "";



            #region Proceso de entregar deliveries.
            if (!fmError && mrCount > 0)
            {

                #region Crear una lista donde no se repita los material request

                var groupedData = wb.AsEnumerable()
                                .GroupBy(row => new
                                {

                                    MaterialRequest = row.Field<string>("MATERIAL_REQUEST"),
                                    Plant = row.Field<string>("PLANT")

                                })
                                    .Where(group => group.Key.MaterialRequest != null && group.Key.MaterialRequest.ToString() != "")
                                    .Select(group => new
                                    {

                                        MaterialRequest = group.Key.MaterialRequest,
                                        Plant = group.Key.Plant,
                                        Materials = group.Select(row => new MaterialData
                                        {
                                            Item = row.Field<string>("ITEM").ToString(),
                                            Material = row.Field<string>("MATERIAL").ToString(),
                                            Quantity = row.Field<string>("QUANTITY"),
                                        }).Distinct().ToList(),

                                        PurchaseOrder = group.Select(row => row.Field<string>("PURCHASE_ORDER")).FirstOrDefault(),
                                        StorageLocation = group.Select(row => row.Field<string>("STORAGE_LOCATION")).FirstOrDefault(),
                                        OrderService = group.Select(row => row.Field<string>("ORDER_SERVICE")).FirstOrDefault(),
                                        CustomerPOOrder = group.Select(row => row.Field<string>("CUSTOMER_PO")).FirstOrDefault(),
                                        Migo = group.Select(row => row.Field<string>("MIGO")).FirstOrDefault(),
                                        PostingDate = group.Select(row => row.Field<string>("POSTING_DATE")).FirstOrDefault()
                                    });


                #endregion

                foreach (var item in groupedData)
                {
                    string po = "", materialRequest/*MR*/ = "", plant = "", storageLocation = "", itemMr = "", material = "", quantity = "", delivery = "", orderService = "", customerPOOrder = "",
                        migo = "", postingDate = "", documentId = "", ot = "", response = ""; // pais_ship = "", respuesta = "";

                    //Asignación de variables.
                    materialRequest = item.MaterialRequest.ToString();
                    plant = item.Plant;

                    po = item.PurchaseOrder.ToString();
                    storageLocation = item.StorageLocation.ToString();
                    orderService = item.OrderService.ToString();
                    customerPOOrder = item.CustomerPOOrder.ToString();
                    migo = item.Migo.ToString();
                    postingDate = item.PostingDate.ToString();

                    bool haveError = false;



                    List<MaterialData> list = item.Materials;

                    #region Verificar si ha sido creado anteriormente en DB.
                    bool created = false;
                    try
                    {
                        string sqlCreate = $"SELECT COUNT(*) as 'quantity' FROM Delivery WHERE materialRequest='{materialRequest}'";
                        created = int.Parse(crud.Select(sqlCreate, "delivery_db", enviroment).Rows[0]["quantity"].ToString()) > 0;
                    }
                    catch (Exception e) { }
                    #endregion

                    //Solo las que terminan en 02 se ejecutan.
                    if (storageLocation != "" && storageLocation.Substring(2, 2) == "02" && !created)
                    {
                        try
                        {
                            #region extraer data y validacion
                            materialRequest = materialRequest.Replace("\n", "");
                            materialRequest = materialRequest.Replace("\r", "");
                            bool Numeric = int.TryParse(materialRequest, out int num);
                            if (Numeric == false)
                            {
                                response = materialRequest + ": " + "el documento no es correspondiente";
                                DataRow row = excelResult.Rows.Add();
                                row["Material Request"] = materialRequest;
                                row["Planta (país)"] = plant;
                                row["Resultado"] = response;
                                continue;
                            }



                            #endregion

                            #region SAP
                            console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                            RfcDestination dest_crm = sap.GetDestRFC(systemSap, mandante);
                            IRfcFunction createDelivery = dest_crm.Repository.CreateFunction("ZFI_DELIVERY_CREATE");
                            createDelivery.SetValue("CONSIGMENT_MR", materialRequest);
                            createDelivery.SetValue("SHIPP_POINT_PAIS", plant);
                            IRfcTable matInfo = createDelivery.GetTable("MATERIAL_INFO");
                            matInfo.Clear();
                            bool cont = true;
                            foreach (MaterialData materialInfo in list)
                            {
                                //por cada material que encuentra 

                                material = materialInfo.Material;
                                quantity = materialInfo.Quantity.Split(',')[0].ToString();
                                itemMr = materialInfo.Item;


                                bool Numeric2 = int.TryParse(quantity, out int num2);
                                if (Numeric2 == false)
                                {
                                    response = materialRequest + ": " + "por favor ingrese una cantidad correcta";
                                    respFinal = respFinal + "\\n" + materialRequest + " : " + response;
                                    cont = false;
                                    DataRow row = excelResult.Rows.Add();
                                    row["Purchase Order"] = po;
                                    row["Material Request"] = materialRequest;
                                    row["Planta (país)"] = plant;
                                    row["Storage Location"] = storageLocation;
                                    row["Item"] = itemMr;
                                    row["WMS Order Transfer"] = ot;
                                    row["Material"] = material;
                                    row["Cantidad"] = quantity;
                                    row["Delivery"] = delivery;
                                    row["Order Service"] = orderService;
                                    row["Customer PO Order"] = customerPOOrder;
                                    row["Migo"] = migo;
                                    row["Posting Date"] = postingDate;
                                    row["ID del Documento"] = documentId;
                                    row["WMS Order Transfer"] = ot;
                                    row["Resultado"] = response;
                                    haveError = true;

                                    //Query de values a insertar a la db.
                                    queryValues += $@"( '{po}', '{plant}', '{storageLocation}', '{materialRequest}', '{itemMr}', '{material}', '{quantity.Split(',')[0]}', '{delivery}', '{orderService}', '{customerPOOrder}', '{migo}', {(postingDate != "0000-00-00" ? "'" + postingDate + "'" : "NULL")}, '{ot}', '{response}', '1','{(haveError == true ? 1 : 0)}' ,'1', 'Databot'),";

                                    break;
                                }


                                matInfo.Append();
                                matInfo.SetValue("MATERIAL", material);
                                matInfo.SetValue("QUANTITY", quantity);
                                matInfo.SetValue("ITEM", itemMr);

                            }

                            if (cont)
                            {
                                createDelivery.Invoke(dest_crm);
                                ot = "";

                                if (createDelivery.GetValue("RESPUESTA").ToString() == "Delivery ya existe")
                                {
                                    response = createDelivery.GetValue("RESPUESTA").ToString();
                                    delivery = createDelivery.GetValue("ID_DOC").ToString();
                                    haveError = true;
                                }
                                else if (createDelivery.GetValue("ID_DOC").ToString() == "")
                                {
                                    response = createDelivery.GetValue("RESPUESTA").ToString();
                                    haveError = true;
                                }
                                else if (createDelivery.GetValue("ID_OT").ToString() == "")
                                {
                                    response = "Se creo la salida, sin embargo no se generó la OT";
                                    delivery = createDelivery.GetValue("ID_DOC").ToString();
                                    haveError = true;
                                }
                                else
                                {
                                    response = "Salida creada con exito, OT: " + createDelivery.GetValue("ID_OT").ToString();
                                    ot = createDelivery.GetValue("ID_OT").ToString();
                                    delivery = createDelivery.GetValue("ID_DOC").ToString();
                                }


                                console.WriteLine(materialRequest + " : " + response);
                                respFinal = respFinal + "\\n" + materialRequest + " : " + response;
                                //log de base de datos
                                log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Delivery", documentId + " : " + response, root.Subject);
                                executeStats = true;

                                //agregar las respuestas al excel
                                foreach (MaterialData materialInfo in list)
                                {
                                    //por cada material que encuentra 
                                    material = materialInfo.Material;
                                    quantity = materialInfo.Quantity.ToString();
                                    itemMr = materialInfo.Item;

                                    DataRow row = excelResult.Rows.Add();

                                    row["Purchase Order"] = po;
                                    row["Material Request"] = materialRequest;
                                    row["Planta (país)"] = plant;
                                    row["Storage Location)"] = storageLocation;
                                    row["Item"] = itemMr;
                                    row["WMS Order Transfer"] = ot;
                                    row["Material"] = material;
                                    row["Cantidad"] = quantity;
                                    row["Delivery"] = delivery;
                                    row["Order Service"] = orderService;
                                    row["Customer PO Order"] = customerPOOrder;
                                    row["Migo"] = migo;
                                    row["Posting Date"] = postingDate;
                                    row["ID del Documento"] = documentId;
                                    row["WMS Order Transfer"] = ot;
                                    row["Resultado"] = response;

                                    //Query de values a insertar a la db.
                                    queryValues += $@"( '{po}', '{plant}', '{storageLocation}', '{materialRequest}', '{itemMr}', '{material}', '{quantity.Split(',')[0]}', '{delivery}', '{orderService}', '{customerPOOrder}', '{migo}', {(postingDate != "0000-00-00" ? "'" + postingDate + "'" : "NULL")}, '{ot}', '{response}', '1','{(haveError == true ? 1 : 0)}' ,'1', 'Databot'),";

                                }


                            }
                            #endregion

                        }
                        catch (Exception ex)
                        {
                            console.WriteLine(ex.Message);
                        }

                    }
                }



                #region Insertar a la db los deliveries creados.
                if (queryValues != "")
                {
                    string query = @" 
                    INSERT INTO Delivery 
                        (purchaseOrder, plant, storageLocation, materialRequest, item, material, quantity, delivery, orderService,
                        customerPOOrder, migo, postingDate, orderTransfer, response, statusReport, haveError, active, createdBy)
                    VALUES  " + (queryValues.Remove(queryValues.Length - 1));
                    console.WriteLine("Insertando deliveries en la db.");
                    bool result = crud.Insert(query, "delivery_db", enviroment);

                    if (!result)
                    {
                        //Si salió mal el insert se guarda en un excel como respaldo.
                        currentDate = DateTime.Now.ToString("yyyy/MM/dd-HH.mm.ss");

                        root.requestDetails = respFinal;
                        excelResult.AcceptChanges();
                        resultPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Databot\\Delivery" + "\\" + $"deliveryResults{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx"; ;
                        xs.CreateExcel(excelResult, "Sheet1", resultPath, true);

                        sett.SendError(this.GetType(), $"No se insertaron los deliveries en la db", $"Se crearon los deliveries, sin embargo, ocurrió un error y no se insertaron en la db para el reporte posterior. Se respaldaron en un excel, y están en la ruta: {resultPath}. Por favor reenviar el excel a adm_inventarios@gbm.net y revisar en los logs porqué no se pudo insertar en la base de datos.");
                        console.WriteLine("Por un error, se guardan deliveries en excel, en la ruta: " + resultPath);
                    }

                    console.WriteLine("Deliveries creados");
                }
                else
                {
                    console.WriteLine($"Se finalizó el proceso de deliveries sin inserción a la db. Habían {mrCount} mrs por crear pero ya existían en la db, por tanto, lo más probable es que esas mrs tenían un error anteriormente entonces por eso no se reprocesaron para no duplicar la data en la db. Si requiere procesar esas mrs nuevamente inactive esas mrs en la db en la tabla de delivery.");
                }
                #endregion

            }
            #endregion


            #region Reportar error.
            if (mrCount > 0)
            {
                if (validar_lineas == false || fmError)
                {
                    console.WriteLine("Reportando errores a administradores.");
                    //enviar email de repuesta de error
                    string[] cc = { "epiedra@gbm.net", "dmeza@gbm.net" };

                    response_failure = "Existe un error en la function module ZFI_GET_MR al extraer los material requests de SAP para los deliveries.";

                    mail.SendHTMLMail(respFinal + "<br>" + response_failure, new string[] {"appmanagement@gbm.net"}, "Error al crear deliveries - " + currentDate, cc);

                }
            }
            else
            {
                console.WriteLine("No existen MRs a procesar");

                #region Insertar a la db los deliveries creados manualmente.
                if (queryValues != "")
                {
                    try
                    {
                        string query = @" 
                    INSERT INTO Delivery 
                        (purchaseOrder, plant, storageLocation, materialRequest, item, material, quantity, delivery, orderService,
                        customerPOOrder, migo, postingDate, orderTransfer, response, statusReport, haveError, active, createdBy)
                    VALUES  " + (queryValues.Remove(queryValues.Length - 1));
                        crud.Insert(query, "delivery_db", enviroment);

                        console.WriteLine("Deliveries manuales insertados a la db.");
                    }
                    catch (Exception e) { }
                }
                #endregion

            }
            #endregion

            #region Reporta en caso que sea la hora planificada.
            if (TimeToReport() || root.BDActivate)
            {
                Report();
            }
            #endregion

        }

        /// <summary>
        ///  Método para validar si es momento de reportar los deliveries a correos de stock.
        /// </summary>
        /// <param ></param>
        /// <returns>True o False si es momento o no a reportar.</returns>
        /// 
        public bool TimeToReport()
        {

            DataTable hours = crud.Select("select hour from HoursToReport where active=1", "delivery_db", enviroment);
            TimeSpan currentHour = new TimeSpan(DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);

            foreach (DataRow hourDr in hours.Rows)
            {
                TimeSpan hour = TimeSpan.Parse(hourDr["hour"].ToString());

                //Se genera un rango de 10 minutos antes y 25 después.
                TimeSpan startRange = hour.Subtract(TimeSpan.FromMinutes(10));
                TimeSpan endRange = hour.Add(TimeSpan.FromMinutes(25));

                //Si está en el rango.
                if (currentHour >= startRange && currentHour <= endRange)
                {
                    console.WriteLine("Momento de reportar.");
                    return true;
                }

            }


            console.WriteLine("No es momento de reportar.");
            return false;
        }

        /// <summary>
        ///  Método para reportar los deliveries pendientes a notificar.
        /// </summary>
        /// <param ></param>
        /// <returns>Vacío.</returns>
        /// 
        public void Report()
        {

            DataTable emails = crud.Select("CALL `getEmails`()", "delivery_db", enviroment);
            DataTable deliveries = crud.Select("CALL `getDeliveries`()", "delivery_db", enviroment); //WHERE statusReport = 1
            int subjectErrors = int.Parse(crud.Select("CALL getErrorsQuantity()", "delivery_db", enviroment).Rows[0]["errors"].ToString());

            if (deliveries.Rows.Count > 0)
            {
                //Generar el update de ids.
                string updateDeliveries = "";

                foreach (DataRow delivery in deliveries.Rows)
                {
                    updateDeliveries += $"'{delivery["id"]}',";
                }

                if (deliveries.Columns.Contains("id"))
                    deliveries.Columns.Remove("id");



                string currentDate2 = DateTime.Now.ToString("yyyy/MM/dd-HH.mm.ss");

                //root.requestDetails = respFinal;
                deliveries.AcceptChanges();
                string resultPath2 = root.FilesDownloadPath + "\\" + $"deliveryResults{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";
                xs.CreateExcel(deliveries, "Sheet1", resultPath2, true);

                string sender = "";
                string[] cc = { };

                //Correos
                foreach (DataRow email in emails.Rows)
                {
                    if (email["type"].ToString() == "1" /*Sender*/)
                    {
                        sender = email["email"].ToString();
                    }
                    else if (email["type"].ToString() == "2"/*cc*/)
                    {
                        cc = cc.Concat(new string[] { email["email"].ToString() }).ToArray();
                    }
                }

                string[] adjunto = { resultPath2 };
                string subject = subjectErrors > 0 ? "Error al crear deliveries - " + currentDate2 : "Deliveries creados exitosamente - " + currentDate2;

                mail.SendHTMLMail("Deliveries creados, los resultados están en el excel.", new string[] { sender }, subject, cc, adjunto);


                //Actualizar status a 2-Reported
                crud.Update($"update Delivery set statusReport=2 where id in ({updateDeliveries.Remove(updateDeliveries.Length - 1)})", "delivery_db", enviroment);

            }
            else
            {
                console.WriteLine("No hay registros a reportar.");
            }
        }

        /// <summary>
        ///  Método para identificar cuales deliveries que fueron creados por otras personas manualmente que deben agregarse a la db. 
        /// </summary>
        /// <param ></param>
        /// <returns>Vacío.</returns>
        /// 
        public string getDeliveriesToAddDataBase(DataTable wb)
        {
            string queryResult = "";

            if (wb.Rows.Count > 0)
            {
                List<string> deliveriesFromSAP = new List<string>();

                #region Listar los deliveries creados de SAP.
                foreach (DataRow row in wb.Rows)
                {
                    string delivery = row["delivery"].ToString();

                    if (delivery != "")
                    {
                        deliveriesFromSAP.Add(delivery);
                    }
                }

                if (deliveriesFromSAP.Count == 0)
                {
                    return queryResult;
                }

                #endregion

                #region Crear query select.
                string querySelect = "select delivery from Delivery where active =1 and delivery in (";

                foreach (string delivery in deliveriesFromSAP)
                {
                    querySelect += $"'{delivery}',";
                }
                querySelect = (querySelect.Remove(querySelect.Length - 1)) + ")";
                #endregion

                DataTable deliveriesFromDBdt = crud.Select(querySelect, "delivery_db", enviroment);

                //Lista de deliveries de la DB que tienen en común con los de SAP.
                List<string> deliveriesFromDB = deliveriesFromDBdt.AsEnumerable().Select(row => row.Field<string>("delivery")).ToList();

                //Lista de deliveries filtrados para agregar a la db. ==> Dinámica: Lista A y Lista B, en la lista A sólo quedan los que no tiene en común con B.
                List<string> deliveriesToAdd = deliveriesFromSAP.Where(item => !deliveriesFromDB.Contains(item)).ToList();

                //DataTable de deliveries para agregar a la db.
                DataTable deliveriesToAddDt =
                    wb.AsEnumerable().Where(row => deliveriesToAdd.Contains(row.Field<string>("delivery")))
                .Any() ? wb.AsEnumerable().Where(row => deliveriesToAdd.Contains(row.Field<string>("delivery"))).CopyToDataTable() : wb.Clone();

                if (deliveriesToAddDt.Rows.Count > 0)
                {
                    console.WriteLine($"Se encontraron {deliveriesToAddDt.Rows.Count} deliveries creados manualmente sin el robot. ");
                }

                foreach (DataRow row in deliveriesToAddDt.Rows)
                {
                    queryResult += $@"( '{row["PURCHASE_ORDER"]}', '{row["PLANT"]}', '{row["STORAGE_LOCATION"]}', '{row["MATERIAL_REQUEST"]}', '{row["ITEM"]}', '{row["MATERIAL"]}', '{row["QUANTITY"].ToString().Split(',')[0]}', '{row["DELIVERY"]}', '{row["ORDER_SERVICE"]}', '{row["CUSTOMER_PO"]}', '{row["MIGO"]}', {(row["POSTING_DATE"].ToString() != "0000-00-00" ? "'" + row["POSTING_DATE"] + "'" : "NULL")}, '{row["ORDER_TRANSFER"]}', '{"Salida creada con exito, OT: " + row["ORDER_TRANSFER"] /*Response*/}', '1',' {'0'/*haveError}*/}' ,'1', 'ADM_INVENTARIOS'),";
                    console.WriteLine($"Mr: {row["MATERIAL_REQUEST"]} con el delivery: {row["DELIVERY"]} en cola para agregarse a la db.");
                }


            }

            return queryResult;
        }
    }
    class MaterialData
    {
        public string Item { get; set; }
        public string Material { get; set; }
        public string Quantity { get; set; }
    }
}

