using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using DataBotV5.Data.Projects.Freelance;
using DataBotV5.Data.Database;
using DataBotV5.Data.SAP;
using DataBotV5.App.Global;
using DataBotV5.Security;
using System.Text;
using DataBotV5.Logical.Encode;
using System.Diagnostics;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.Webex;

namespace DataBotV5.Logical.Projects.Freelance
{
    /// <summary>
    /// Clase Logical encargada de Freelance FI.
    /// </summary>
    class FreelanceFiSS : IDisposable
    {
        private bool disposedValue;
        SapVariants sap = new SapVariants();
        CRUD crud = new CRUD();
        ConsoleFormat console = new ConsoleFormat();
        SecureAccess sec = new SecureAccess();
        FreelanceSqlSS freelanceSql = new FreelanceSqlSS();
        MailInteraction mail = new MailInteraction();
        string syst = "ERP";
        string ssMandante = "PRD";

        //public billEmails DeterminateOrganization(string po)
        //{
        //    Dictionary<string, string> parameters = new Dictionary<string, string>();
        //    parameters["PO"] = po;

        //    IRfcFunction func = new SapVariants().ExecuteRFC(mandante, "ZCATS_INFO_PO_CTY", parameters);

        //    string pais = func.GetValue("CTY").ToString();

        //    DataTable mytable = new DataTable();
        //    string sql = $"SELECT * FROM freelance_mf WHERE PAIS = '{pais}'";
        //    mytable = new CRUD().Select("Databot", sql, "automation");

        //    string contadores = $"{mytable.Rows[0][2]}";
        //    string copias = $"{mytable.Rows[0][3]}";

        //    List<string> correos = new List<string>();
        //    if (contadores.Contains(","))
        //    {
        //        string[] cnt = contadores.Split(',');
        //        correos = cnt.ToList();
        //    }
        //    else
        //    {
        //        correos.Add(contadores);
        //    }


        //    List<IFreelanceBase> freelanceBase = new List<IFreelanceBase>();
        //    InfoFI ifi = new InfoFI { Emails = correos, Copias = copias };
        //    freelanceBase.Add(ifi);
        //    return freelanceBase;
        //}
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cats"></param>
        /// <param name="po"></param>
        /// <param name="item"></param>
        /// <param name="id"></param>
        /// <param name="consultant"></param>
        /// <param name="consultantUser"></param>
        /// <param name="consultantEmail"></param>
        /// <param name="responsable"></param>
        public bool CreateSheet(DataRow hesRow)
        {
            string cats = hesRow["catsId"].ToString();
            string po = hesRow["purchaseOrder"].ToString();
            string item = hesRow["item"].ToString();
            string id = hesRow["hesId"].ToString();
            //string consultant = hesRow["consultant"].ToString();
            //string consultantName = hesRow["consultantName"].ToString();
            string consultantEmail = hesRow["consultantEmail"].ToString();
            string responsable = hesRow["responsible"].ToString();
            string byHito = hesRow["byHito"].ToString();
            createHES HES = new createHES();
            try
            {
                HES = ProcessSheet(po, item, byHito, HES); //CREA LA HES MEDIANTE LA TRANSACCION CATM
                if (byHito == "0")
                {
                    //solo se puede extraer cuando NO es por hito
                    HES.HES = ExtractSheetCATS(cats); //BUSCA LA HES MEDIANTE EL NUMERO DE CATS QUE SE GENERO CUANDO SE APROBARON LAS HORAS
                }
                if (HES.HES == "" || HES.HES == null)
                {
                    //Fallo creando la hoja, notifique
                    //MySqlConnection conn = new Database().Conn("automation");
                    string imagen64 = "";
                    using (BinaryFiles bf = new BinaryFiles())
                    {
                        //bf.DataBaseInsertOrUpdate($"UPDATE freelance_hoja SET ESTADO = 'ERROR', SCREENSHOT = @bina WHERE ID = '{id}'", conn, "bina", HES.Captura);
                        imagen64 = bf.BinaryHTMLImage(HES.Captura);
                    }

                    freelanceSql.updateErrorHes(id, HES.Captura, po, item, "2", consultantEmail.Split('@')[0], ssMandante);


                    string mensaje = $"Estimado coordinador, la HES de la PO:{po}, Item:{item} no ha podido ser creada por el siguiente error: </br>{imagen64}";
                    string subject = $"Portal Freelance: Error de HES (Creacion)";
                    string html = Properties.Resources.emailtemplate1;
                    html = html.Replace("{subject}", "Portal Proveedores: error al crear la HES");
                    html = html.Replace("{cuerpo}", mensaje);
                    html = html.Replace("{contenido}", "");

                    //mail.SendNotificationErrorSheet($"", mensaje, "Portal Freelance: Error de HES (Creacion)");
                    mail.SendHTMLMail(html, new string[] { "hogonzalez@gbm.net" }, subject, new string[] { "dmeza@gbm.net", $"{responsable}@GBM.NET" }, null);
                    return false;
                }
                else
                {
                    string mensaje = "";
                    string imagen64 = "";

                    //Todo bien Libere, notifique al Coordinador y Freelance para que adjunte las facturas
                    freelanceSql.updateHes(id, HES.Captura, po, item, HES.HES, ssMandante);

                    FreeHes les = new FreeHes();
                    if (byHito == "0")
                    {
                        //Solo se hace cuando NO es por hitos
                        les = FreeSheet(HES.HES);
                    }
                    else
                    {
                        les.HES = HES.HES;
                        les.status = true;
                    }
                    if (les.status)
                    {
                        //true todo salio bien

                    }
                    else
                    {

                        using (BinaryFiles bf = new BinaryFiles())
                        {
                            //bf.DataBaseInsertOrUpdate($"UPDATE freelance_hoja SET ESTADO = 'CREADA', FECHA = '{DateTime.Now.ToString("dd.MM.yyyy")}' ,SCREENSHOT = @bina WHERE HES = '{HES}'", conn, "bina", les.Captura);
                            imagen64 = bf.BinaryHTMLImage(les.screenShot);
                        }
                        freelanceSql.updateErrorHes(id, les.screenShot, po, item, "4", consultantEmail.Split('@')[0], ssMandante);

                        //Algo salio mal, notifique al coordinador el error para que lo libere manualmente
                        mensaje = $"Estimado coordinador, la HES {les.HES} no ha sido liberada en SAP satisfactoriamente, favor realizar los ajustes manuales en SAP del siguiente error: </br>{imagen64}";
                        string subject = $"Portal Freelance: Error de liberacion de HES";
                        string html = Properties.Resources.emailtemplate1;
                        html = html.Replace("{subject}", "Portal Proveedores: Error de liberacion de HES");
                        html = html.Replace("{cuerpo}", mensaje);
                        html = html.Replace("{contenido}", "");
                        mail.SendHTMLMail(html, new string[] { "hogonzalez@gbm.net" }, subject, new string[] { "dmeza@gbm.net", $"{responsable}@GBM.NET" }, null);
                        //return false; //no devuelve falso si da error a la hora de liberar la HES
                        //mail.SendNotificationErrorSheet($"{responsable}@GBM.NET", mensaje, "Portal Freelance: Error de liberacion de HES");
                    }

                }
                return true;
            }
            catch (Exception exe)
            {
                string mensaje = $"Estimado coordinador, dio error al intentar crear la HES de la PO {po} - {item},  </br> {exe.Message} </br> {exe.StackTrace}";
                string subject = $"Portal Freelance: Error al crear/liberar la HES";
                string html = Properties.Resources.emailtemplate1;
                html = html.Replace("{subject}", "Portal Proveedores: Error inesperado de HES");
                html = html.Replace("{cuerpo}", mensaje + $"{responsable}@GBM.NET");
                html = html.Replace("{contenido}", "");
                mail.SendHTMLMail(html, new string[] { "hogonzalez@gbm.net" }, subject, new string[] { "dmeza@gbm.net" }, null);
                return false;
            }




        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cats"></param>
        /// <returns></returns>
        private string ExtractSheetCATS(string cats)
        {
            string HES = "";
            //List<string> hojas = new List<string>();
            //List<CATSRecords> cats = new List<CATSRecords>();
            //DataTable registros = crud.Select("Databot", $"SELECT ID,CATS FROM freelance_g WHERE PO = '{po}' AND ITEM = '{item}' AND ESTADO = 'APROBADO' AND HOJA = '0'", "automation");
            //if (registros != null)
            //{
            //    if (registros.Rows.Count > 0)
            //    {
            //        for (int i = 0; i < registros.Rows.Count; i++)
            //        {
            //            CATSRecords cATS = new CATSRecords
            //            {
            //                Id = $"{registros.Rows[i][0]}",
            //                Records = $"{registros.Rows[i][1]}"
            //            };
            //            cats.Add(cATS);
            //        }
            //    }
            //}
            //cats.ForEach(x =>
            //{
            HES = CatsH(cats); //AQUI CREA LA HP HOJA
            //});

            //hojas = cats.Select(x => x.HES).Distinct().ToList();
            //if (hojas.Count == 1)
            //{
            //    HES = hojas[0];
            //    //Hay solo una hoja, actualice los registros

            //    cats.ForEach(x =>
            //    {
            //        string sql = $"UPDATE freelance_g SET HOJA = '{HES}' WHERE ID = '{x.Id}'";
            //        crud.Update("Databot", sql, "automation");
            //    });
            //}


            return HES;
        }

        private createHES ProcessSheet(string po, string item, string byHito, createHES HES)
        {
            try
            {

                if (item.Length == 4)
                {
                    item = $"0{item}";
                }
                if (byHito == "1")
                {
                    ((SAPFEWSELib.GuiFrameWindow)SapVariants.session.FindById("wnd[0]")).Maximize();
                    ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nml81n";
                    SapVariants.frame.SendVKey(0);
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[17]")).Press();
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtRM11R-EBELN")).Text = po;
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/txtRM11R-EBELP")).Text = item;
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[13]")).Press();
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txtESSR-TXZ01")).Text = po;
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/subSERVICE:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW/ctxtESLL-SRVPOS[2,0]")).Text = "SAP_CONSULTORIA";
                    SapVariants.frame.SendVKey(0);
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[25]")).Press();
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[11]")).Press(); //guardar
                    try
                    {

                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/usr/btnSPOP-OPTION1")).Press();
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[2]/tbar[0]/btn[0]")).Press(); //ventana de error
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtIMKPF-BLDAT")).Text = "10.04.2023";
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtIMKPF-BUDAT")).Text = "10.04.2023";
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();

                    }
                    catch (Exception)
                    {

                    }
                    try
                    {
                        HES.HES = ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtESSR-LBLNI")).Text;
                    }
                    catch (Exception)
                    {
                        HES.HES = "";
                    }

                    byte[] img = ((SAPFEWSELib.GuiFrameWindow)SapVariants.session.ActiveWindow).HardCopyToMemory(2);
                    HES.Po = po;
                    HES.Item = item;
                    HES.Captura = img;


                }
                else
                {
                    ((SAPFEWSELib.GuiFrameWindow)SapVariants.session.FindById("wnd[0]")).Maximize();
                    ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/ncatm";
                    SapVariants.frame.SendVKey(0);
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtCATSEKKO-SEBELN")).Text = po;
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtCATSEKKO-SEBELN")).SetFocus();
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();

                    string mensaje = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text;


                    if (mensaje == "")
                    {
                        try
                        {
                            for (int i = 4; i < 200; i++)
                            {
                                try
                                {
                                    string objectPo = ((SAPFEWSELib.GuiLabel)SapVariants.session.FindById($"wnd[1]/usr/lbl[6,{i}]")).Text;
                                    string objectItem = ((SAPFEWSELib.GuiLabel)SapVariants.session.FindById($"wnd[1]/usr/lbl[18,{i}]")).Text;

                                    if (objectPo == po && objectItem == item)
                                    {
                                        ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById($"wnd[1]/usr/chk[2,{i}]")).Selected = true;
                                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[7]")).Press();
                                        //tome captura

                                        byte[] img = ((SAPFEWSELib.GuiFrameWindow)SapVariants.session.ActiveWindow).HardCopyToMemory(2);
                                        HES.Po = po;
                                        HES.Item = item;
                                        HES.Captura = img;
                                        SapVariants.frame.SendVKey(0);
                                        break;
                                    }
                                    else
                                    {
                                        using (WebexTeams wb = new WebexTeams())
                                        {
                                            wb.SendNotification("dmeza@gbm.net", "Error coincidencia de items", $"El item {item} de la Po {po} no coincide con el de SAP: {objectItem}");
                                        }
                                    }
                                }
                                catch (Exception)
                                { }

                            }
                        }
                        catch (Exception)
                        {

                            //((SAPFEWSELib.GuiFrameWindow)SAP_Variants.session.FindById("wnd[0]")).Maximize();
                            byte[] img = ((SAPFEWSELib.GuiFrameWindow)SapVariants.session.ActiveWindow).HardCopyToMemory(2);
                            HES.Po = po;
                            HES.Item = item;
                            HES.Captura = img;

                        }
                    }
                    else
                    {
                        byte[] img = ((SAPFEWSELib.GuiFrameWindow)SapVariants.session.ActiveWindow).HardCopyToMemory(2);
                        //No hay nada que hacer en la CATM
                        HES.Po = po;
                        HES.Item = item;
                        HES.Captura = img;
                    }
                }

            }
            catch (Exception ex)
            {
                byte[] img = ((SAPFEWSELib.GuiFrameWindow)SapVariants.session.ActiveWindow).HardCopyToMemory(2);
                //No hay nada que hacer en la CATM
                HES.Po = po;
                HES.Item = item;
                HES.Captura = img;
            }

            return HES;
        }
        private FreeHes FreeSheet(string HES)
        {
            FreeHes les = new FreeHes();

            using (SapVariants sap = new SapVariants())
            {
                try
                {


                    //MySqlConnection conn = new Database().Conn("automation");
                    //sap.LogSAP("300");
                    ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nml81n";
                    SapVariants.frame.SendVKey(0);
                    // ((SAPFEWSELib.GuiContainerShell)SAP_Variants.session.FindById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]")).Top = "          4";
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[17]")).Press();
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtRM11R-LBLNI")).Text = HES;
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[5]")).Press();
                    ((SAPFEWSELib.GuiTab)SapVariants.session.FindById("wnd[0]/usr/tabsTAB_HEADER/tabpREGA")).Select();
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/tabsTAB_HEADER/tabpREGA/ssubSUB_ACCEPTANCE:SAPLMLSR:0420/ctxtESSR-BUDAT")).Text = DateTime.Now.ToString("dd.MM.yyyy");
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[25]")).Press();
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/usr/btnSPOP-OPTION1")).Press();
                    string mensaje = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text;
                    if (!mensaje.Contains("Service entry sheet saved, acceptance document"))
                    {
                        //No se pudo liberar envie mensaje al coordinador, para que lo libere manual
                        les.status = false;
                        les.HES = HES;
                        les.screenShot = ((SAPFEWSELib.GuiFrameWindow)SapVariants.session.ActiveWindow).HardCopyToMemory(2);
                        //using (BinaryFiles bf = new BinaryFiles())
                        //{
                        //    bf.DataBaseInsertOrUpdate($"UPDATE freelance_hoja SET ESTADO = 'ERRORL', SCREENSHOT = @bina, FECHA = '{DateTime.Now.ToString("dd.MM.yyyy")}' WHERE HES = '{HES}'", conn, "bina", les.Captura);
                        //}
                    }
                    else
                    {
                        les.status = true;
                        les.HES = HES;
                        les.screenShot = ((SAPFEWSELib.GuiFrameWindow)SapVariants.session.ActiveWindow).HardCopyToMemory(2);
                        //using (BinaryFiles bf = new BinaryFiles())
                        //{
                        //    bf.DataBaseInsertOrUpdate($"UPDATE freelance_hoja SET ESTADO = 'CREADA', FECHA = '{DateTime.Now.ToString("dd.MM.yyyy")}' ,SCREENSHOT = @bina WHERE HES = '{HES}'", conn, "bina", les.Captura);
                        //}
                        //Liberado exitosamente comunique al coordinador y freelance para que adjunte las facturas
                    }
                }
                catch (Exception ex)
                {
                    console.WriteLine(ex.Message);
                    les.status = false;
                    les.HES = HES;
                    les.screenShot = ((SAPFEWSELib.GuiFrameWindow)SapVariants.session.ActiveWindow).HardCopyToMemory(2);
                }
            }
            return les;
        }
        //private void GenerateSheetAndFree()
        //{
        //    DataTable hojas = freelanceSql.GetxSheet();
        //    for (int i = 0; i < hojas.Rows.Count; i++)
        //    {
        //        string id = hojas.Rows[i][0].ToString();
        //        string po = hojas.Rows[i][1].ToString();
        //        string item = hojas.Rows[i][2].ToString();
        //        string tiempo = hojas.Rows[i][5].ToString();

        //        DateTime par = DateTime.Parse(tiempo);
        //        string fecha = par.Day.ToString("00") + "." + par.Month.ToString("00") + "." + par.Year.ToString();

        //        Process.Start(@"C:\Users\databot02\Desktop\Freelance.lnk");
        //        System.Threading.Thread.Sleep(8000);

        //        //EasyLDR con = new EasyLDR();
        //        string item_hoja;
        //        if (item.Length <= 4)
        //        {
        //            item_hoja = "0" + item;
        //        }
        //        else
        //        {
        //            item_hoja = item;
        //        }
        //        //con.HOJA(po, item_hoja);
        //        string hoja_up = "";

        //        //DataTable registros = freelance.GetxCats(po, item);
        //        //for (int x = 0; x < registros.Rows.Count; x++)
        //        //{
        //        //    string cats = registros.Rows[x][0].ToString();
        //        //    string hoja = registros.Rows[x][1].ToString();

        //        //    if (cats != "" && hoja == "0")
        //        //    {
        //        //        string v_hoja = CatsH(cats);
        //        //        if (v_hoja != "")
        //        //        {
        //        //            hoja_up = v_hoja;
        //        //            string update_query_g = "UPDATE freelance_g SET HOJA = '" + hoja_up + "' WHERE CATS = '" + cats + "'";
        //        //            crud.Update("Databot", update_query_g, "automation");

        //        //        }
        //        //    }
        //        //}

        //        //string update_query = "UPDATE freelance_h SET SAP = 'D', HOJA = '" + hoja_up + "' WHERE ID = '" + id + "'";
        //        //crud.Update("Databot", update_query, "automation");

        //        //con.LHOJA(hoja_up, fecha);

        //        //process.KillProcess("saplogon", false);
        //    }
        //}
        //private IRfcTable TablaLongTextM(IRfcFunction func, string texto)
        //{
        //    IRfcTable longt = func.GetTable("LGTXT");

        //    string textolargo = sec.DecodePass(texto);
        //    textolargo = textolargo.Replace("\r\n", string.Empty);
        //    textolargo = textolargo.Replace("\n", string.Empty);
        //    textolargo = textolargo.Replace("\t", string.Empty);



        //    double partSize = textolargo.Length / 130;
        //    int partes = Int32.Parse(partSize.ToString());

        //    if (partes == 0 && textolargo.Length <= 130)
        //    {
        //        longt.Append();
        //        longt.SetValue("ROW", "1");
        //        longt.SetValue("FORMAT_COL", "*");
        //        longt.SetValue("TEXT_LINE", textolargo);
        //    }
        //    else if (partes > 0)
        //    {
        //        int inicio = 0;
        //        int fin = 130;

        //        if (partes > 7)
        //        {
        //            partes = 7;
        //        }
        //        for (int i = 0; i < partes; i++)
        //        {
        //            longt.Append();
        //            longt.SetValue("ROW", "1");
        //            if (i == 0)
        //            {
        //                longt.SetValue("FORMAT_COL", "*");

        //                longt.SetValue("TEXT_LINE", textolargo.Substring(inicio, fin));
        //            }
        //            else
        //            {
        //                longt.SetValue("FORMAT_COL", "*");
        //                try
        //                {
        //                    longt.SetValue("TEXT_LINE", textolargo.Substring(inicio, fin));
        //                }
        //                catch (Exception ex)
        //                {
        //                    int finc = i * 130;
        //                    int contador = (textolargo.Substring(inicio, textolargo.Length - finc)).Length;
        //                    longt.SetValue("TEXT_LINE", textolargo.Substring(inicio, textolargo.Length - finc));
        //                }

        //            }
        //            inicio += 130;
        //            fin += 130;
        //        }
        //    }
        //    else
        //    {
        //        longt.Append();
        //        longt.SetValue("ROW", "1");
        //        longt.SetValue("FORMAT_COL", "*");
        //        longt.SetValue("TEXT_LINE", "");
        //    }

        //    return longt;
        //}
        /// <summary>
        /// 
        /// </summary>
        /// <param name="func"></param>
        /// <param name="texto"></param>
        /// <returns></returns>
        private IRfcTable TablaLongText(IRfcFunction func, string texto)
        {
            IRfcTable longt = func.GetTable("LONGTEXT");

            string textolargo = texto;
            textolargo = textolargo.Replace("\r\n", string.Empty);
            textolargo = textolargo.Replace("\n", string.Empty);
            textolargo = textolargo.Replace("\t", string.Empty);



            double partSize = textolargo.Length / 130;
            int partes = Int32.Parse(partSize.ToString());

            if (partes == 0 && textolargo.Length <= 130)
            {
                longt.Append();
                longt.SetValue("ROW", "1");
                longt.SetValue("FORMAT_COL", "*");
                longt.SetValue("TEXT_LINE", textolargo);
            }
            else if (partes > 0)
            {
                int inicio = 0;
                int fin = 130;

                if (partes > 7)
                {
                    partes = 7;
                }
                for (int i = 0; i < partes; i++)
                {
                    longt.Append();
                    longt.SetValue("ROW", "1");
                    if (i == 0)
                    {
                        longt.SetValue("FORMAT_COL", "*");

                        longt.SetValue("TEXT_LINE", textolargo.Substring(inicio, fin));
                    }
                    else
                    {
                        longt.SetValue("FORMAT_COL", "*");
                        try
                        {
                            longt.SetValue("TEXT_LINE", textolargo.Substring(inicio, fin));
                        }
                        catch (Exception ex)
                        {
                            int finc = i * 130;
                            int contador = (textolargo.Substring(inicio, textolargo.Length - finc)).Length;
                            longt.SetValue("TEXT_LINE", textolargo.Substring(inicio, textolargo.Length - finc));
                        }

                    }
                    inicio += 130;
                    fin += 130;
                }
            }
            else
            {
                longt.Append();
                longt.SetValue("ROW", "1");
                longt.SetValue("FORMAT_COL", "*");
                longt.SetValue("TEXT_LINE", "");
            }
            //}

            return longt;
        }
        static IEnumerable<string> WholeChunks(string str, int chunkSize)
        {
            for (int i = 0; i < str.Length; i += chunkSize)
                yield return str.Substring(i, chunkSize);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="counter"></param>
        /// <returns></returns>
        private string CatsH(string counter)
        {
            string cats = "";
            console.WriteLine(" Conectado con SAP ERP");
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters["COUNTER"] = counter;

            IRfcFunction func = sap.ExecuteRFC(syst, "ZCATS_HOJA", parameters);


            string respuesta = func.GetValue("RESPONSE").ToString();
            if (respuesta != "N/A")
            {
                cats = respuesta;
            }
            return cats;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="hours">horas</param>
        /// <param name="profile">tipo</param>
        /// <param name="counter">cats id</param>
        /// <param name="ticket">ticket</param>
        /// <param name="detail">detalles</param>
        /// <returns></returns>
        public string CatsM(string hours, string profile, string counter, string ticket, string detail)
        {
            string respuesta = "";
            try
            {


                if (hours.Contains("."))
                {
                    hours = hours.Replace(".", ",");
                }
                Dictionary<string, string> parameters = new Dictionary<string, string>();
                parameters["COUNTER"] = counter;
                parameters["PROFILE"] = profile;
                parameters["HOURS"] = hours;
                parameters["TICKET"] = ticket;

                IRfcFunction func = sap.ExecuteRFC(syst, "ZCATS_EDITAR_REPORTE", parameters);

                //IRfcTable longtext = TablaLongTextM(func, detail);

                respuesta = func.GetValue("RESPONSE").ToString();
                //if (respuesta == "OK")
                //{
                //    cats = respuesta;
                //}

            }
            catch (Exception ex)
            {
                respuesta = ex.Message;
            }
            return respuesta;
        }
        /// <summary>
        /// Metodo que crea los CATS en SAP
        /// </summary>
        /// <param name="po"></param>
        /// <param name="it"></param>
        /// <param name="hours"></param>
        /// <param name="date"></param>
        /// <param name="employee"></param>
        /// <param name="detail"></param>
        /// <param name="ticket"></param>
        /// <param name="profile"></param>
        /// <returns></returns>
        public CatInfo Cats(string po, string it, string hours, string date, string employee, string detail, string ticket, string profile)
        {
            CatInfo cat = new CatInfo();
            try
            {
                RfcDestination destination = sap.GetDestRFC(syst);
                RfcRepository repo = destination.Repository;
                IRfcFunction func = repo.CreateFunction("ZCATS_REPORTE");
                IRfcTable general = func.GetTable("GENERAL_DATA");
                IRfcTable longtext = TablaLongText(func, detail);

                if (hours.Contains(","))
                {
                    hours = hours.Replace(",", ".");
                }


                general.Append();
                general.SetValue("PO", po);
                general.SetValue("ITEM", it);
                general.SetValue("HOURS", hours);
                general.SetValue("DATE_R", DateTime.Parse(date).ToString("yyyy-MM-dd"));
                general.SetValue("FREELANCE", employee);
                general.SetValue("TICKET", ticket);

                func.SetValue("PROFILE", profile);

                func.Invoke(destination);


                string respuesta = func.GetValue("RESPONSE").ToString();
                if (respuesta == "OK")
                {
                    cat.CatId = func.GetValue("CATS").ToString();
                    cat.RespError = null;

                }
                else
                {
                    string error_repo = "";
                    IRfcTable errores = func.GetTable("SAP_RESPONSE");
                    for (int i = 0; i < errores.Count; i++)
                    {
                        var x = errores.CurrentIndex = i;
                        error_repo = error_repo + "\r\n" + errores.CurrentRow[3].GetValue().ToString();
                    }

                    string respuestasap = error_repo;
                    console.WriteLine(respuestasap);
                    cat.CatId = null;
                    cat.RespError = respuestasap;
                }
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                cat.CatId = null;
                cat.RespError = ex.ToString();

            }

            return cat;
        }

        /// <summary>
        /// Metodo para aprobar el cats
        /// </summary>
        /// <param name="tipo"></param>
        /// <param name="cats"></param>
        /// <param name="reason"></param>
        /// <param name="profile"></param>
        /// <returns></returns>
        public string CatsAppr(string tipo, string cats, string reason, string profile)
        {
            string respuesta = "";

            console.WriteLine(" Conectado con SAP ERP");

            Dictionary<string, string> parameters = new Dictionary<string, string>
            {
                ["CATS"] = cats,
                ["STATUS_C"] = tipo,
                ["REASON"] = reason,
                ["PROFILE"] = profile
            };

            IRfcFunction func = sap.ExecuteRFC(syst, "ZCATS_F_APROBAR", parameters);

            respuesta = func.GetValue("RESPONSE").ToString();

            return respuesta;
        }

        /// <summary>
        /// Método para obtener el nombre de una tabla de facturación.
        /// </summary>
        /// <param name="object">Po asociadas a la factura</param>
        /// <param name="provider">persona que subio la factura</param>
        /// <returns></returns>
        public string BillingTable(DataTable poInfo, string provider)
        {
            poInfo.DefaultView.Sort = "purchaseOrder ASC";
            poInfo = poInfo.DefaultView.ToTable();
            StringBuilder strHTMLBuilder = new StringBuilder();
            strHTMLBuilder.Append("<table class='myCustomTable' width='100 %'>");
            strHTMLBuilder.Append("<thead>");
            strHTMLBuilder.Append("<tr>");
            strHTMLBuilder.Append("<th>Proveedor</th>");
            strHTMLBuilder.Append("<th>Purchase Order</th>");
            strHTMLBuilder.Append("<th>Item</th>");
            strHTMLBuilder.Append("<th>HES</th>");
            strHTMLBuilder.Append($"<th>{((poInfo.Rows[0]["byHito"].ToString() == "1") ? "Monto" : "Horas")}</th>");
            strHTMLBuilder.Append("</thead>");
            strHTMLBuilder.Append("<tbody>");
            double totalHours = 0;
            for (int i = 0; i < poInfo.Rows.Count; i++)
            {

                if (i == 0)
                {
                    double hoursActual = double.Parse(poInfo.Rows[i]["hours"].ToString());
                    totalHours = totalHours + hoursActual;
                    if (i == poInfo.Rows.Count - 1)
                    {
                        strHTMLBuilder.Append($"<tr><td class='tg-zv4m'>{provider}</td><td class='tg-zv4m'>{poInfo.Rows[i]["purchaseOrder"].ToString()}</td><td class='tg-zv4m'>{poInfo.Rows[i]["item"].ToString()}</td><td class='tg-zv4m'>{poInfo.Rows[i]["hesNumber"].ToString()}</td><td class='tg-zv4m'>{((poInfo.Rows[0]["byHito"].ToString() == "1") ? poInfo.Rows[i]["mountHito"].ToString() : poInfo.Rows[i]["hours"].ToString())}</td></tr>");
                        totalHours = 0;
                    }
                }
                else
                {

                    //string poActual = poInfo.Rows[i]["purchaseOrder"].ToString();
                    //string poAnterior = poInfo.Rows[i - 1]["purchaseOrder"].ToString();
                    string hesActual = poInfo.Rows[i]["hesNumber"].ToString();
                    string hesAnterior = poInfo.Rows[i - 1]["hesNumber"].ToString();
                    if (hesActual == hesAnterior)
                    {
                        //int hoursActual = int.Parse(poInfo.Rows[i]["hours"].ToString());
                        //int hoursAnterior = int.Parse(poInfo.Rows[i - 1]["hours"].ToString());
                        double hoursActual = double.Parse(poInfo.Rows[i]["hours"].ToString());
                        double hoursAnterior = double.Parse(poInfo.Rows[i - 1]["hours"].ToString());
                        totalHours = totalHours + hoursActual;
                        //var totalHours = hoursActual + hoursAnterior;
                        //identificar si es la ultima iteracion
                        if (i == poInfo.Rows.Count - 1)
                        {
                            strHTMLBuilder.Append($"<tr><td class='tg-zv4m'>{provider}</td><td class='tg-zv4m'>{poInfo.Rows[i]["purchaseOrder"].ToString()}</td><td class='tg-zv4m'>{poInfo.Rows[i]["item"].ToString()}</td><td class='tg-zv4m'>{poInfo.Rows[i]["hesNumber"].ToString()}</td><td class='tg-zv4m'>{((poInfo.Rows[0]["byHito"].ToString() == "1") ? poInfo.Rows[i]["mountHito"].ToString() : totalHours.ToString())}</td></tr>");
                            totalHours = 0;
                        }

                    }
                    else
                    {
                        strHTMLBuilder.Append($"<tr><td class='tg-zv4m'>{provider}</td><td class='tg-zv4m'>{poInfo.Rows[i - 1]["purchaseOrder"].ToString()}</td><td class='tg-zv4m'>{poInfo.Rows[i - 1]["item"].ToString()}</td><td class='tg-zv4m'>{poInfo.Rows[i - 1]["hesNumber"].ToString()}</td><td class='tg-zv4m'>{((poInfo.Rows[0]["byHito"].ToString() == "1") ? poInfo.Rows[i]["mountHito"].ToString() : totalHours.ToString())}</td></tr>");
                        totalHours = 0;
                        double hoursActual = double.Parse(poInfo.Rows[i]["hours"].ToString());
                        double hoursAnterior = double.Parse(poInfo.Rows[i - 1]["hours"].ToString());
                        totalHours = totalHours + hoursActual;

                        if (i == poInfo.Rows.Count - 1)
                        {
                            strHTMLBuilder.Append($"<tr><td class='tg-zv4m'>{provider}</td><td class='tg-zv4m'>{poInfo.Rows[i]["purchaseOrder"].ToString()}</td><td class='tg-zv4m'>{poInfo.Rows[i]["item"].ToString()}</td><td class='tg-zv4m'>{poInfo.Rows[i]["hesNumber"].ToString()}</td><td class='tg-zv4m'>{((poInfo.Rows[0]["byHito"].ToString() == "1") ? poInfo.Rows[i]["mountHito"].ToString() : poInfo.Rows[i]["hours"].ToString())}</td></tr>");
                            totalHours = 0;
                        }
                    }


                }

            }
            strHTMLBuilder.Append("</tbody>");
            strHTMLBuilder.Append("</table>");
            string Htmltext = strHTMLBuilder.ToString();
            return Htmltext;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~FreelanceFI()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        void IDisposable.Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
    //public class InfoFI : IFreelanceBase
    //{
    //    private string _copias;
    //    public List<string> Emails { get; set; }
    //    public string Copias { get { return _copias; } set { _copias = value; } }

    //}
    public class FreeHes
    {
        public bool status { get; set; }
        public string HES { get; set; }
        public byte[] screenShot { get; set; }
    }
    public class createHES
    {
        public string HES { get; set; }
        public string Po { get; set; }
        public string Item { get; set; }
        public byte[] Captura { get; set; }
    }
    public class billEmails
    {
        public string[] senders { get; set; }
        public string[] copies { get; set; }
    }

    public class CatInfo
    {
        public string CatId { get; set; }
        public string RespError { get; set; }
    }
}
