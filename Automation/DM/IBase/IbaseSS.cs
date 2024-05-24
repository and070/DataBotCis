using System;
using Newtonsoft.Json.Linq;
using SAP.Middleware.Connector;
using DataBotV5.Data.Projects.MasterData;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using DataBotV5.Logical.Webex;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using DataBotV5.Data.SAP;

namespace DataBotV5.Automation.DM.Ibase

{
    /// <summary>
    /// Clase DM Automation encargada de IBase.
    /// </summary>
    class IbaseSS
    {
        ConsoleFormat console = new ConsoleFormat();
        MailInteraction mail = new MailInteraction();
        Rooting root = new Rooting();
        Log log = new Log();
        MasterDataSqlSS DM = new MasterDataSqlSS();
        WebexTeams wt = new WebexTeams();
        Stats stats = new Stats();

        string description, ibaseName, ibaseCountry, ibaseCustomer, inputType, equips, response, responseFailure, ibaseID, action, resp = "", message = "";
        bool err = false;
        string[] equipsList;
        string crmMand = "CRM";

        string respFinal = "";


        public void Main()
        {
            string respuesta = DM.GetManagement("5"); //IBASE
            if (!String.IsNullOrEmpty(respuesta) && respuesta != "ERROR")
            {
                console.WriteLine("Procesando...");
                ProcessIBase();
                console.WriteLine("Creando Estadísticas");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }

        }
        public void ProcessIBase()
        {
            try
            {
                root.requestDetails = root.requestDetails.Replace("\u00A0", " "); //eliminar non breaks spaces (char 160)
                root.requestDetails = root.requestDetails.Replace(@"[^\u0000-\u007F]+", ""); //eliminar caracteres no ASCII

                if (root.metodoDM == "1") //lineal
                {
                    JArray requests = JArray.Parse(root.requestDetails);
                    string requestType = root.tipo_gestion; //1 crear , 2 modificar

                    for (int i = 0; i < requests.Count; i++)
                    {
                        JObject fila = JObject.Parse(requests[i].ToString());
                        string functionModuleName = "ZDM_IB_CRE";

                        switch (requestType)
                        {
                            case "1"://nueva
                                description = fila["description"].Value<string>().ToUpper();
                                ibaseName = fila["name"].Value<string>().ToUpper();
                                ibaseCountry = fila["gbmCountriesCode"].Value<string>().ToUpper();
                                ibaseCustomer = fila["client"].Value<string>();
                                equips = fila["equipments"].Value<string>();
                                inputType = fila["incomeMethodCode"].Value<string>();
                                resp = "Creación";
                                break;
                            case "2"://modificación
                                ibaseID = fila["ibase"].Value<string>().ToUpper();
                                action = fila["actionCode"].Value<string>().ToUpper();
                                equips = fila["equipment"].Value<string>();
                                inputType = fila["incomeMethodCode"].Value<string>();
                                if (action == "ADD")
                                    functionModuleName = "ZDM_IB_MOD";

                                resp = "Modificación";
                                break;
                            default:
                                break;
                        }

                        equipsList = equips.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);

                        #region SAP
                        console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                        try
                        {
                            RfcDestination destination = new SapVariants().GetDestRFC(crmMand);
                            IRfcFunction functionModule = destination.Repository.CreateFunction(functionModuleName);
                            IRfcTable equipsTable = functionModule.GetTable("EQUIPOS");
                            IRfcTable equipsTableOut = functionModule.GetTable("EQUIPOS_OUT");

                            #region Parámetros de SAP
                            switch (requestType)
                            {
                                case "1":
                                    functionModule.SetValue("DESCRIPCION", description);
                                    functionModule.SetValue("NOMBRE", ibaseName);
                                    functionModule.SetValue("PAIS", ibaseCountry);
                                    functionModule.SetValue("CLIENTE", ibaseCustomer);
                                    break;
                                case "2":
                                    switch (action)
                                    {
                                        case "ADD":
                                            functionModule.SetValue("IBASE", ibaseID);
                                            break;
                                        case "DEL":
                                            functionModule.SetValue("DESCRIPCION", "DEL_" + ibaseID);
                                            functionModule.SetValue("NOMBRE", "DELETED_IBASE");
                                            break;
                                        default:
                                            break;
                                    }
                                    break;
                                default:
                                    break;
                            }

                            foreach (string equip in equipsList)
                            {
                                equipsTable.Append();
                                if (inputType == "ID")
                                {
                                    try
                                    {
                                        if (10000000 < int.Parse(equip) && int.Parse(equip) < 19999999)
                                        {
                                            equipsTable.SetValue("R3IDENT", equip);
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        message = "Se colocó un ID no valido";
                                        break;
                                    }

                                }
                                else
                                    equipsTable.SetValue("R3SER_NO", equip.ToUpper());
                            }

                            #endregion
                            if (message == "")
                            {
                                #region Invocar FM
                                functionModule.Invoke(destination);
                                #endregion
                                #region Procesar Salidas del FM
                                message = functionModule.GetValue("MENSAJE").ToString();
                                if (requestType == "1") //crear
                                    ibaseID = functionModule.GetValue("IBASE").ToString();

                                err = ibaseID == "" ? true : false;
                                for (int j = 0; j < equipsTableOut.RowCount; j++)
                                {
                                    err = (equipsTableOut[j].GetValue("MESSAGE").ToString() == "Diferentes equipos" || equipsTableOut[j].GetValue("MESSAGE").ToString() == "ID duplicado" || equipsTableOut[j].GetValue("MESSAGE").ToString() == "No se encontró el equipo") ? true : false;
                                    equipsList[j] = "<td>" + equipsTableOut[j].GetValue("R3IDENT").ToString().TrimStart('0') + "</td>" + "<td>" + equipsTableOut[j].GetValue("R3SER_NO").ToString().TrimStart('0') + "</td>" + "<td>" + equipsTableOut[j].GetValue("MESSAGE").ToString() + "</td>";
                                    equipsList[j] = "<tr>" + equipsList[j] + "</tr>";

                                    string response = "Se agregaron los siguientes equipos: " + equipsTableOut[j].GetValue("R3IDENT").ToString().TrimStart('0') + ", " + equipsTableOut[j].GetValue("R3SER_NO").ToString().TrimStart('0') + ", "+ equipsTableOut[j].GetValue("MESSAGE").ToString() ;
                                    log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear equipos Ibase", response, "");

                                    respFinal = respFinal + "\\n" + "Se agregaron los siguientes equipos: " + response;

                                }
                                string temp = string.Join("", equipsList);

                                if (action == "ADD")
                                    response = message + "<br>Se agregaron los siguientes equipos<br>" + "<table border=0><tr><th>ID</th><th>Serie</th><th>Mensaje</th></tr>" + temp + "</table><br>" + response;
                                else if (action == "DEL")
                                    response = message + "<br>Se eliminaron los siguientes equipos<br>" + "<table border=0><tr><th>ID</th><th>Serie</th><th>Mensaje</th></tr>" + temp + "</table><br>" + response;
                                else
                                    response = message + "<br>" + "<table border=0><tr><th>ID</th><th>Serie</th><th>Mensaje</th></tr>" + temp + "</table><br>" + response;

                                response = "<br>" + resp + " del ibase: " + ibaseID.TrimStart('0') + "<br>" + response + "<br>";

                                #endregion
                            }

                            root.requestDetails = respFinal;
                        }
                        catch (Exception ex)
                        {
                            responseFailure = new ValidateData().LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, i);
                            console.WriteLine("Finishing process " + responseFailure);
                            response = response + ibaseName + ": " + ex.ToString() + "<br>";
                            responseFailure = ex.ToString();
                            DM.ChangeStateDM(root.IdGestionDM, "Ibase: " + ibaseID + "<br>" + "" + "<br>", "4"); //ERROR
                        }
                        #endregion

                    }
                }

                console.WriteLine("Finalizando solicitud");
                if (err == true)
                {
                    //error pero no finalizar
                    string[] cc = { "smarin@gbm.net" };
                    DM.ChangeStateDM(root.IdGestionDM, "Ibase: " + ibaseID + "<br>" + response + "<br>", "4"); //ERROR
                    mail.SendHTMLMail("Gestión: " + root.IdGestionDM + "<br>" + "Error realizando la " + resp + " del ibase: " + ibaseID + "<br>" + response + "<br>", new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, cc);
                }
                else
                {
                    //log.LogDeCambios(resp, root.BDProcess, new string[] { root.BDUserCreatedBy }, resp + " Ibase", ibaseID, root.Subject);
                    DM.ChangeStateDM(root.IdGestionDM, "Se realizo la " + resp + " del ibase: " + ibaseID, "3"); //FINALIZADO
                    if (message != "")
                    {
                        if (message == "Se coloco un ID no valido")
                        {
                            DM.ChangeStateDM(root.IdGestionDM, message, "4"); //ERROR
                            string resp_end = "**Notificación de gestión de Ibase:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha finalizado, con el siguiente resultado: <br><br> " + message + "<br>";

                            mail.SendHTMLMail(resp_end, new string[] { root.BDUserCreatedBy }, "Error: " + root.Subject);
                        }
                        else
                        {
                            //enviar email de repuesta de error a datos maestros
                            string[] cc = { "smarin@gbm.net" };
                            mail.SendHTMLMail("Gestión: " + root.IdGestionDM + "<br>" + response + "<br>", new string[] { root.BDUserCreatedBy }, "Error: " + root.Subject, cc);
                        }
                    }
                    else
                    {
                        //finalizar solicitud
                        mail.SendHTMLMail("Gestión: " + root.IdGestionDM + "<br>" + response + "<br>", new string[] { root.BDUserCreatedBy }, root.Subject);

                        string resp_chat = "**Notificación de gestión de Ibase:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha finalizado, con el siguiente resultado (ver log completo en el email que acaba de llegarle): <br><br> " + "Se realizo la " + resp;
                        wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, resp_chat);
                    }
                }
            }
            catch (Exception ex)
            {
                string[] cc = { "smarin@gbm.net" };
                DM.ChangeStateDM(root.IdGestionDM, ex.Message, "4"); //ERROR
                mail.SendHTMLMail("Gestión: " + root.IdGestionDM + "<br>" + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, cc);
            }
        }
    }
}
