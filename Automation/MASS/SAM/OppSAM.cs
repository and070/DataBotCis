using Newtonsoft.Json.Linq;
using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Root;
using DataBotV5.Data.Database;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Projects.OppSAM;
using DataBotV5.App.Global;
using DataBotV5.Data.Projects.OppSAM;
using DataBotV5.Logical.Webex;

namespace DataBotV5.Automation.MASS.SAM
{
    /// <summary>
    /// Clase MASS Automation encargada de gestión de oportunidades en SAM.
    /// </summary>
    class OppSAM
    {
        
        ConsoleFormat console = new ConsoleFormat();
        OppSAMSQL opp = new OppSAMSQL();
        MailInteraction mail = new MailInteraction();
        Credentials cred = new Credentials();
        Rooting root = new Rooting();
        MQListener MQ = new MQListener();
        Stats estadisticas = new Stats();
        SapVariants sap = new SapVariants();
        CRUD crud = new CRUD();
        Log log = new Log();
        string mandante = "CRM";
        WebexTeams wx = new WebexTeams();
        string respFinal = "";
        bool executeStats = false;



        public void Main()
        {
            string[] res_opp = null;
            MQ.listener();

            DirectoryInfo directory = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Databot\\OPP_SAM\\");
            foreach (FileInfo file in directory.GetFiles())
            {
                try
                {
                    root.Mq_mensaje[1] = File.ReadAllText(file.FullName);

                    if (root.Mq_mensaje[1].Contains("ERROR"))
                    {
                        int intentos = Int32.Parse(root.Mq_mensaje[2]); ;
                        //Error, intentar reconeccion
                        if (intentos > 5)
                        {
                            mail.SendHTMLMail(root.Mq_mensaje[1] + "<br><br>Se intento reconectar pero fallo", new string[] { "internalcustomersrvs@gbm.net" }, "Se perdio la conexión con el MQ");
                            opp.TurnOffRobot();
                            root.Mq_mensaje[0] = "0";
                            root.Mq_mensaje[1] = "";
                            root.Mq_mensaje[2] = "0";
                        }
                        else
                        {
                            intentos++;
                            root.Mq_mensaje[2] = intentos.ToString();
                            root.Mq_mensaje[0] = "0";
                        }
                    }
                    else
                    {
                        //Todo correcto...
                        console.WriteLine("Procesando y notificando");

                        List<string> equipos_list = new List<string>();

                        JObject json_in = JObject.Parse(root.Mq_mensaje[1]);
                        string contacto = json_in["correoUsuario"].Value<string>().ToUpper();
                        string cliente = json_in["idCliente"].Value<string>().ToUpper();
                        string nombre_contacto = json_in["nombreUsuario"].Value<string>().ToUpper();
                        string nombre_cliente = json_in["nombreCliente"].Value<string>().ToUpper();
                        JArray equipos = (JArray)json_in["equipos"];
                        string fecha_fin = "";

                        foreach (JObject element in equipos)
                        {
                            equipos_list.Add(element["numeroSerie"].Value<string>());
                            fecha_fin = element["fechaVencimientoGarantia"].Value<string>();
                        }


                        //GUI
                        //bool check_login = sap.CheckLogin(mandante);
                        //if (!check_login)
                        //{
                        //    sap.BlockUser(mandante, 1);
                        //RFC
                        res_opp = ProcessOPP(contacto, cliente, equipos_list, fecha_fin, false); //{RESPONSE, OPP_ID, PAIS, TERRITORIO, EMAIL}

                        //    sap.BlockUser(mandante, 0);
                        //}
                        //else
                        //{
                        //    res_opp[0] = "GUI ocupado";
                        //}


                        //PARA PRUEBAS///////////////////////////////////////////////////////////////////////////////////////////
                        //res_opp[4] = "smarin@GBM.NET";                                                                       //
                        //string[] res_opp = { "", "000000777", "CR", "002", "smarin@GBM.NET" };//ejemplo de respuesta ok////////


                        if (res_opp[0] != "" || res_opp[1] == "")
                        {
                            string[] ata = { file.FullName };
                            mail.SendHTMLMail("Error FM: " + res_opp[0] + "<br><br>JSON:" + root.Mq_mensaje[1], new string[] { "internalcustomersrvs@gbm.net" }, "Error: OPP SAM", attachments: ata);
                            //notificar a la gente de SAM??????

                        }
                        else
                        {
                            executeStats = true;
                            //Enviar Notificaciones  //Cuidado con enviar un SPAM terrible
                            string[] usuarios = Notify(res_opp[2]); //{nombre INFRA,infra_email,nombre CSM,csm_email}
                            string[] territorio = opp.GetManager(res_opp[2], res_opp[3]); //{nombre, correo}
                            string[] senders = new string[1];
                            string mensaje;
                            string subject = "Oportunidad " + res_opp[1].TrimStart('0') + " para soporte de HW";

                            //CSM
                            mensaje = EmailFormat(usuarios[2], nombre_contacto, nombre_cliente, res_opp[1].TrimStart('0'), equipos_list);
                            senders[0] = usuarios[3];
                            if (senders[0] != "" && senders[0] != null)
                            {
                                mail.SendHTMLMail(mensaje, senders , subject);
                                wx.SendNotification(usuarios[3], "OPP", mensaje);
                            }

                            //INFRA
                            mensaje = EmailFormat(usuarios[0], nombre_contacto, nombre_cliente, res_opp[1].TrimStart('0'), equipos_list);
                            senders[0] = usuarios[1];
                            if (senders[0] != "" && senders[0] != null)
                            {
                                mail.SendHTMLMail(mensaje, senders, subject);
                                wx.SendNotification(usuarios[1], "OPP", mensaje);
                            }

                            //Responsable
                            mensaje = EmailFormat("", nombre_contacto, nombre_cliente, res_opp[1].TrimStart('0'), equipos_list);
                            senders[0] = res_opp[4];
                            if (senders[0] != "" && senders[0] != null)
                            {
                                mail.SendHTMLMail(mensaje, senders , subject);
                                wx.SendNotification(res_opp[4], "OPP", mensaje);
                            }

                            //Gerente Ventas
                            mensaje = EmailFormat(territorio[0], nombre_contacto, nombre_cliente, res_opp[1].TrimStart('0'), equipos_list);
                            senders[0] = territorio[1];
                            if (senders[0] != "" && senders[0] != null)
                            {
                                mail.SendHTMLMail(mensaje, senders , subject);
                                wx.SendNotification(territorio[1], "OPP", mensaje);
                            }

                            //Contacto
                            senders[0] = contacto;
                            if (senders[0] != "" && senders[0] != null)
                            {
                                mail.SendHTMLMail(EmailFormat("", nombre_contacto, nombre_cliente, res_opp[1].TrimStart('0'), equipos_list), senders , subject);
                            }
                            //guardar_solicitud(root.Mq_mensaje[1], res_opp[1].TrimStart('0')); // para guardar la opp en otra tabla, al final la vi innecesaria
                            log.LogDeCambios("Creacion", root.BDProcess, "SAM(IBM MQ)", "Oportunidad: " + res_opp[1].TrimStart('0'), root.Mq_mensaje[1], res_opp[0]);
                            respFinal = respFinal + "\\n" + "Oportunidad: " + res_opp[1].TrimStart('0') + " " + root.Mq_mensaje[1] + ", " + res_opp[0];


                            //root.Mq_mensaje[1] = "";
                        }


                    }
                }

                catch (Exception ex)
                {
                    string[] ata = { file.FullName };
                    if (res_opp != null)
                    {
                        log.LogDeCambios("Creacion", root.BDProcess, "SAM(IBM MQ)", "Oportunidad: " + res_opp[1].TrimStart('0'), root.Mq_mensaje[1], "ERROR");
                        respFinal = respFinal + "\\n" + "Error Oportunidad: " + res_opp[1].TrimStart('0') + " " + root.Mq_mensaje[1] + ", " + res_opp[0];

                        mail.SendHTMLMail("Error FM: " + res_opp[0] + "<br><br>JSON:" + root.Mq_mensaje[1] + "<br><br>Error: " + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, "Error: OPP SAM", attachments: ata);
                    }
                    else
                    {
                        log.LogDeCambios("Creacion", root.BDProcess, "SAM(IBM MQ)", "Oportunidad: no se creo", root.Mq_mensaje[1], "ERROR");
                        respFinal = respFinal + "\\n" + "Error Oportunidad no se creó: " + res_opp[1].TrimStart('0') + " " + root.Mq_mensaje[1] + ", " + res_opp[0];

                        mail.SendHTMLMail("Error FM: no se llamó<br><br>JSON:" + root.Mq_mensaje[1] + "<br><br>Error: " + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, "Error: OPP SAM", attachments: ata);
                    }
                }
                file.Delete();
            }

            if (executeStats == true)
            {

                root.requestDetails = respFinal;
                root.BDUserCreatedBy = "KVILLAR";

                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }
        public string[] ProcessOPP(string contact, string client, List<string> teams, string DateEnd, bool GUI)
        {
            string response_failure;
            string[] resultado = new string[5];
            #region SAP
            console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
            try
            {
                RfcDestination destination = sap.GetDestRFC(mandante);
                RfcRepository repo = destination.Repository;
                IRfcFunction func = repo.CreateFunction("ZRPA_SAM_OPP_CREATE");
                IRfcTable fm_equipos = func.GetTable("EQUIPOS");
                IRfcFunction func2 = repo.CreateFunction("ZOPP_EQUI_WRITE");
                IRfcTable fm_equipos2 = func2.GetTable("EQUIPOS");
                string equipos_notas = "";

                #region Parametros de SAP
                func.SetValue("CONTACTO", contact.ToLower());
                func.SetValue("CLIENTE", client);
                func.SetValue("FECHA_FIN", DateEnd);
                func.SetValue("TIPO", ""); //"X" = opp por pais  // "" = ZOPS o comentar
                if (GUI == true)
                {
                    func.SetValue("GET_DATA", "X");
                }

                for (int j = 0; j < teams.Count; j++)
                {
                    fm_equipos.Append();
                    fm_equipos.SetValue("TDLINE", teams[j]);
                    equipos_notas = equipos_notas + teams[j] + System.Environment.NewLine;
                }
                #endregion

                #region Invocar FM
                func.Invoke(destination);
                #endregion
                #region Procesar Salidas del FM
                resultado[0] = func.GetValue("RESPONSE").ToString();
                resultado[1] = func.GetValue("OPP_ID").ToString();
                resultado[2] = func.GetValue("PAIS").ToString();
                resultado[3] = func.GetValue("TERRITORIO").ToString();
                resultado[4] = func.GetValue("EMAIL").ToString();

                #region Si GUI se activa
                if (GUI == true)//Procesar los datos para GUI
                {
                    IRfcTable general_data = func.GetTable("GENERAL_DATA");
                    IRfcTable partners_data = func.GetTable("PARTNERS_DATA");
                    IDictionary<string, string> partners_data_list = new Dictionary<string, string>();
                    string[] partner_data = new string[2];

                    string TIPO = general_data[0].GetValue("TIPO").ToString();
                    string DESCRIPCION = general_data[0].GetValue("DESCRIPCION").ToString();
                    string STRING_INICIO = general_data[0].GetValue("FECHA_INICIO").ToString();
                    string STRING_FIN = general_data[0].GetValue("FECHA_FIN").ToString();
                    string FASE_VENTAS = general_data[0].GetValue("FASE_VENTAS").ToString();
                    string PORCENTAJE = general_data[0].GetValue("PORCENTAJE").ToString();
                    string ORIGEN = general_data[0].GetValue("ORIGEN").ToString();
                    string PRIORIDAD = general_data[0].GetValue("PRIORIDAD").ToString();
                    string SRV_ORG_DATA = func.GetValue("SRV_ORG_DATA").ToString();
                    string SALES_ORG_DATA = func.GetValue("SALES_ORG_DATA").ToString();
                    SALES_ORG_DATA = SALES_ORG_DATA.Replace("O ", "");

                    DateTime FECHA_INICIO = Convert.ToDateTime(STRING_INICIO);
                    DateTime FECHA_FIN = Convert.ToDateTime(STRING_FIN);

                    for (int j = 0; j < partners_data.RowCount; j++)
                    {
                        partners_data_list.Add(partners_data[j].GetValue("FUNCTION").ToString(), partners_data[j].GetValue("PARTNER").ToString());
                    }

                    SapVariants sap = new SapVariants();
                    //sap.MatarProceso("saplogon",false);
                    sap.LogSAP(mandante.ToString());

                    //SAP_Variants.frame.Iconify();
                    ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/ncrmd_order";
                    SapVariants.frame.SendVKey(0);
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[9]")).Press();
                    ((SAPFEWSELib.GuiToolbarControl)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_COMMON:SAPLCRM_1O_UI:3150/subSCR_1O_TBAR_CREATE:SAPLCRM_1O_UI:7160/cntlCONTAINER_1O_TBAR0/shellcont/shell")).PressContextButton("1OMAIN_CREATE");
                    ((SAPFEWSELib.GuiToolbarControl)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_COMMON:SAPLCRM_1O_UI:3150/subSCR_1O_TBAR_CREATE:SAPLCRM_1O_UI:7160/cntlCONTAINER_1O_TBAR0/shellcont/shell")).SelectContextMenuItem("BUS2000111@" + TIPO + "@1OMAIN_CREATE");
                    //details
                    ((SAPFEWSELib.GuiTab)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD01")).Select();
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD01/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:7102/subSCRAREA2:SAPLCRM_OPPORT_UI:7110/subSCRAREA1:SAPLCRM_OPPORT_UI:7111/ctxtCRMT_7110_OPPORT_UI-STARTDATE")).Text = FECHA_INICIO.Day.ToString() + "." + FECHA_INICIO.Month.ToString() + "." + FECHA_INICIO.Year.ToString();
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD01/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:7102/subSCRAREA2:SAPLCRM_OPPORT_UI:7110/subSCRAREA1:SAPLCRM_OPPORT_UI:7111/ctxtCRMT_7110_OPPORT_UI-EXPECT_END")).Text = FECHA_FIN.Day.ToString() + "." + FECHA_FIN.Month.ToString() + "." + FECHA_FIN.Year.ToString();
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD01/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:7102/subSCRAREA1:SAPLCRM_OPPORT_UI:7105/subHEADER_SUBSCREEN:SAPLCRM_OPPORT_UI:7010/txtCRMT_7010_OPPORT_UI-DESCRIPTION")).Text = DESCRIPCION;
                    ((SAPFEWSELib.GuiComboBox)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD01/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:7102/subSCRAREA2:SAPLCRM_OPPORT_UI:7110/subSCRAREA1:SAPLCRM_OPPORT_UI:7111/cmbCRMT_7110_OPPORT_UI-CURR_PHASE")).Key = FASE_VENTAS;
                    ((SAPFEWSELib.GuiComboBox)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD01/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:7102/subSCRAREA4:SAPLCRM_OPPORT_UI:7125/subSCRAREA1:SAPLCRM_OPPORT_UI:7126/cmbCRMT_7120_OPPORT_UI-SOURCE")).Key = ORIGEN;
                    ((SAPFEWSELib.GuiComboBox)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD01/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:7102/subSCRAREA4:SAPLCRM_OPPORT_UI:7125/subSCRAREA1:SAPLCRM_OPPORT_UI:7126/cmbCRMT_7120_OPPORT_UI-IMPORTANCE")).Key = PRIORIDAD;

                    //organization
                    ((SAPFEWSELib.GuiTab)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD15")).Select();
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD15/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:7205/subOPPORT_ORGMAN_SUBSCREEN:SAPLCRM_ORGMAN_UI:1022/subORGMAN_HEADER:SAPLCRM_ORGMAN_UI:1012/subORGMAN_SUBSCREEN_FRAME_HEADER:SAPLCRM_ORGMAN_UI:1003/subORG_DATA_SCREEN_SERVICE_H:SAPLCRM_ORGMAN_UI:1004/ctxtCRMT_1004_ORGMAN_UI-SERVICE_ORG_RESP_SHORT")).Text = SRV_ORG_DATA;
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD15/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:7205/subOPPORT_ORGMAN_SUBSCREEN:SAPLCRM_ORGMAN_UI:1022/subORGMAN_HEADER:SAPLCRM_ORGMAN_UI:1012/subORGMAN_SUBSCREEN_FRAME_HEADER:SAPLCRM_ORGMAN_UI:1003/subORG_DATA_SCREEN_SERVICE_H:SAPLCRM_ORGMAN_UI:1004/ctxtCRMT_1004_ORGMAN_UI-SERVICE_ORG_SHORT")).Text = SRV_ORG_DATA;
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD15/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:7205/subOPPORT_ORGMAN_SUBSCREEN:SAPLCRM_ORGMAN_UI:1022/subORGMAN_HEADER:SAPLCRM_ORGMAN_UI:1012/subORGMAN_SUBSCREEN_FRAME_HEADER:SAPLCRM_ORGMAN_UI:1003/subORG_DATA_SCREEN_SALES_H:SAPLCRM_ORGMAN_UI:1001/ctxtCRMT_1001_ORGMAN_UI-SALES_ORG_RESP_SHORT")).Text = SALES_ORG_DATA;
                    SapVariants.frame.SendVKey(0);
                    ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[1]/usr/cntlPFUNC_1021/shellcont/shell/shellcont[1]/shell[1]")).ChangeCheckbox("          2", "&Hierarchy", true);
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();

                    //notas
                    ((SAPFEWSELib.GuiTab)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD09")).Select();
                    ((SAPFEWSELib.GuiTextedit)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD09/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:1210/subTEXT_SUBSCREEN:SAPLCOM_TEXT_MAINTENANCE:2130/subSCRAREA:SAPLCOM_TEXT_MAINTENANCE:2131/cntlSPLITTER_CONTAINER_2131/shellcont/shellcont/shell/shellcont[1]/shell")).Text = equipos_notas;

                    //partner 122975
                    ((SAPFEWSELib.GuiTab)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD06")).Select();
                    ((SAPFEWSELib.GuiComboBox)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD06/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:1180/subPARTNER_SUBSCREEN:SAPLCOM_PARTNER_UI2:2000/subGS_SUBSCREEN_AREA_2000:SAPLCOM_PARTNER_UI2:2050/subGS_SUBSCREEN_AREA_2050:SAPLCOM_PARTNER_UI2:2002/tblSAPLCOM_PARTNER_UI2OVERVIEW_2002/cmbGS_DYNP_2000_PARTNER-PARTNER_FCT[0,0]")).Key = "00000015";
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD06/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:1180/subPARTNER_SUBSCREEN:SAPLCOM_PARTNER_UI2:2000/subGS_SUBSCREEN_AREA_2000:SAPLCOM_PARTNER_UI2:2050/subGS_SUBSCREEN_AREA_2050:SAPLCOM_PARTNER_UI2:2002/tblSAPLCOM_PARTNER_UI2OVERVIEW_2002/ctxtGS_DYNP_2000_PARTNER-PARTNER_NUMBER[1,0]")).Text = partners_data_list["00000015"];
                    ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD06/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:1180/subPARTNER_SUBSCREEN:SAPLCOM_PARTNER_UI2:2000/subGS_SUBSCREEN_AREA_2000:SAPLCOM_PARTNER_UI2:2050/subGS_SUBSCREEN_AREA_2050:SAPLCOM_PARTNER_UI2:2002/tblSAPLCOM_PARTNER_UI2OVERVIEW_2002/chkGS_DYNP_2000_PARTNER-MAINPARTNER[2,0]")).Selected = true;
                    ((SAPFEWSELib.GuiComboBox)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD06/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:1180/subPARTNER_SUBSCREEN:SAPLCOM_PARTNER_UI2:2000/subGS_SUBSCREEN_AREA_2000:SAPLCOM_PARTNER_UI2:2050/subGS_SUBSCREEN_AREA_2050:SAPLCOM_PARTNER_UI2:2002/tblSAPLCOM_PARTNER_UI2OVERVIEW_2002/cmbGS_DYNP_2000_PARTNER-PARTNER_FCT[0,1]")).Key = "00000021";
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD06/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:1180/subPARTNER_SUBSCREEN:SAPLCOM_PARTNER_UI2:2000/subGS_SUBSCREEN_AREA_2000:SAPLCOM_PARTNER_UI2:2050/subGS_SUBSCREEN_AREA_2050:SAPLCOM_PARTNER_UI2:2002/tblSAPLCOM_PARTNER_UI2OVERVIEW_2002/ctxtGS_DYNP_2000_PARTNER-PARTNER_NUMBER[1,1]")).Text = partners_data_list["00000021"];
                    ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD06/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:1180/subPARTNER_SUBSCREEN:SAPLCOM_PARTNER_UI2:2000/subGS_SUBSCREEN_AREA_2000:SAPLCOM_PARTNER_UI2:2050/subGS_SUBSCREEN_AREA_2050:SAPLCOM_PARTNER_UI2:2002/tblSAPLCOM_PARTNER_UI2OVERVIEW_2002/chkGS_DYNP_2000_PARTNER-MAINPARTNER[2,1]")).Selected = true;

                    //sales team
                    ((SAPFEWSELib.GuiTab)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD05")).Select();
                    ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD05/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:1190/subPARTNER_SUBSCREEN:SAPLCOM_PARTNER_UI2:2000/subGS_SUBSCREEN_AREA_2000:SAPLCOM_PARTNER_UI2:2050/subGS_SUBSCREEN_AREA_2050:SAPLCOM_PARTNER_UI2:2002/tblSAPLCOM_PARTNER_UI2OVERVIEW_2002/chkGS_DYNP_2000_PARTNER-MAINPARTNER[2,0]")).Selected = true;
                    ((SAPFEWSELib.GuiComboBox)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD05/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:1190/subPARTNER_SUBSCREEN:SAPLCOM_PARTNER_UI2:2000/subGS_SUBSCREEN_AREA_2000:SAPLCOM_PARTNER_UI2:2050/subGS_SUBSCREEN_AREA_2050:SAPLCOM_PARTNER_UI2:2002/tblSAPLCOM_PARTNER_UI2OVERVIEW_2002/cmbGS_DYNP_2000_PARTNER-PARTNER_FCT[0,0]")).Key = "00000014";
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD05/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:1190/subPARTNER_SUBSCREEN:SAPLCOM_PARTNER_UI2:2000/subGS_SUBSCREEN_AREA_2000:SAPLCOM_PARTNER_UI2:2050/subGS_SUBSCREEN_AREA_2050:SAPLCOM_PARTNER_UI2:2002/tblSAPLCOM_PARTNER_UI2OVERVIEW_2002/ctxtGS_DYNP_2000_PARTNER-PARTNER_NUMBER[1,0]")).Text = partners_data_list["00000014"];

                    ((SAPFEWSELib.GuiTableControl)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD05/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:1190/subPARTNER_SUBSCREEN:SAPLCOM_PARTNER_UI2:2000/subGS_SUBSCREEN_AREA_2000:SAPLCOM_PARTNER_UI2:2050/subGS_SUBSCREEN_AREA_2050:SAPLCOM_PARTNER_UI2:2002/tblSAPLCOM_PARTNER_UI2OVERVIEW_2002")).GetAbsoluteRow(1).Selected = true; //fila 2
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD05/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:1190/subPARTNER_SUBSCREEN:SAPLCOM_PARTNER_UI2:2000/subGS_SUBSCREEN_AREA_2000:SAPLCOM_PARTNER_UI2:2050/btnOVERVIEW_2002_DELETE")).Press();

                    //((SAPFEWSELib.GuiCheckBox)SAP_Variants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD05/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:1190/subPARTNER_SUBSCREEN:SAPLCOM_PARTNER_UI2:2000/subGS_SUBSCREEN_AREA_2000:SAPLCOM_PARTNER_UI2:2050/subGS_SUBSCREEN_AREA_2050:SAPLCOM_PARTNER_UI2:2002/tblSAPLCOM_PARTNER_UI2OVERVIEW_2002/chkGS_DYNP_2000_PARTNER-MAINPARTNER[2,1]")).Selected = true;
                    //((SAPFEWSELib.GuiComboBox)SAP_Variants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD05/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:1190/subPARTNER_SUBSCREEN:SAPLCOM_PARTNER_UI2:2000/subGS_SUBSCREEN_AREA_2000:SAPLCOM_PARTNER_UI2:2050/subGS_SUBSCREEN_AREA_2050:SAPLCOM_PARTNER_UI2:2002/tblSAPLCOM_PARTNER_UI2OVERVIEW_2002/cmbGS_DYNP_2000_PARTNER-PARTNER_FCT[0,1]")).Key = "00000012";
                    //((SAPFEWSELib.GuiTextField)SAP_Variants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\\TOPP_HD05/ssubTABSTRIP_SUBSCREEN:SAPLCRM_OPPORT_UI:1190/subPARTNER_SUBSCREEN:SAPLCOM_PARTNER_UI2:2000/subGS_SUBSCREEN_AREA_2000:SAPLCOM_PARTNER_UI2:2050/subGS_SUBSCREEN_AREA_2050:SAPLCOM_PARTNER_UI2:2002/tblSAPLCOM_PARTNER_UI2OVERVIEW_2002/ctxtGS_DYNP_2000_PARTNER-PARTNER_NUMBER[1,1]")).Text = partners_data_list["00000014"];////el 12?????

                    //save
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();
                    resultado[1] = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text;
                    resultado[1] = resultado[1].Replace("Transaction ", "");
                    resultado[1] = resultado[1].Replace(" saved", "");

                    //guardar en tabla
                    func2.SetValue("OPP_ID", resultado[1]);
                    for (int j = 0; j < teams.Count; j++)
                    {
                        fm_equipos2.Append();
                        fm_equipos2.SetValue("TDLINE", teams[j]);
                    }
                    func2.Invoke(destination); //1457   
                    string res2 = func.GetValue("RESPONSE").ToString();
                    if (res2.Contains("error"))
                    {
                        resultado[0] = func.GetValue("RESPONSE").ToString();
                    }

                }
                #endregion

                #endregion

            }
            catch (Exception ex)
            {
                response_failure = ex.Message;
                console.WriteLine(" Finishing process " + response_failure);
                resultado[0] = response_failure;
                sap.KillSAP();
                sap.BlockUser(mandante, 0);
            }
            return resultado;
            #endregion
        }
        public string[] Notify(string country)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select;
            DataTable mytable = new DataTable();
            string[] usuarios = new string[3];
            try
            {
                #region Connection DB
                sql_select = "select * from notifications where country = '" + country + "'";
                mytable = crud.Select(sql_select, "opp_sam");
                #endregion

                string[] arr = new string[mytable.Columns.Count - 1];
                if (mytable.Rows.Count > 1)
                {
                    usuarios[0] = "Error: Se encontro mas de una fila en un solo pais";
                    return usuarios;
                }
                else
                {
                    arr[0] = mytable.Rows[0][1].ToString();
                    arr[1] = mytable.Rows[0][2].ToString();
                    arr[2] = mytable.Rows[0][3].ToString();
                    arr[3] = mytable.Rows[0][4].ToString();
                }
                usuarios = arr;
            }
            catch (Exception ex)
            {
                usuarios[0] = "Error: " + ex.ToString();
                return usuarios;
            }
            return usuarios;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="personToNotify"></param>
        /// <param name="contact"></param>
        /// <param name="clientName"></param>
        /// <param name="oportunityNumber"></param>
        /// <param name="teams"></param>
        /// <returns></returns>
        public string EmailFormat(string personToNotify, string contact, string clientName, string oportunityNumber, List<string> teams)
        {
            string equi = "";
            for (int j = 0; j < teams.Count; j++)
            {
                equi = equi + teams[j] + ",";
            }
            equi = equi.Remove(equi.Length - 1);

            string template = "<p>Hola <b>" + personToNotify + "<o:p></o:p></b></span></i></p><p><o:p>&nbsp;</o:p></span></i></p><p>El usuario <b>" + contact + " </b>perteneciente al<b> </b>cliente<b> " + clientName +
            " </b>ha solicitado recibir una propuesta de soporte sobre el(los) equipo(s) <b>" + equi + " </b>que posee(n) garantía(s) próxima(s) a vencer.<b> <o:p></o:p></b></span></i></p><p><o:p>&nbsp;</o:p></span></i></p><p>La Oportunidad <b>" + oportunityNumber +
            "</b> fue creada para definir una propuesta de soporte. Por favor revise el detalle de la oportunidad creada y comuníquese con el cliente darle seguimiento a su solicitud. <o:p></o:p></span></i></p>";
            return template;
        }

    }
}
