using System;
using Excel = Microsoft.Office.Interop.Excel;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Projects.SpecialBidForm;
using DataBotV5.Data.Database;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using DataBotV5.Logical.Projects.SpecialBidForms;
using DataBotV5.Logical.Mail;

namespace DataBotV5.Automation.RPA2.SpecialBid
{
    class OldSpecialBidForm
    {
        SpecialBidFormSQL spbidsql = new SpecialBidFormSQL();
        Rooting roots = new Rooting();
        Database db2 = new Database();
        ValidateData val = new ValidateData();
        ConsoleFormat console = new ConsoleFormat();
        SbForm sb = new SbForm();
        ProcessInteraction proc = new ProcessInteraction();
        CRUD crud = new CRUD();
        MailInteraction mail = new MailInteraction();
        Log log = new Log();
        Stats estadisticas = new Stats();
        Credentials cred = new Credentials();
        string sb_number = "";
        public string id_carpeta = "";
        public string projectname;
        public string ibm_country = "";
        public string useopp = "";
        public string opp = "";
        public string useprevbid = "";
        public string prevbid = "";
        public string priceupd = "";
        public string justi = "";
        public string customer = "";
        public string brand = "";
        public string justi2 = "";
        public string addquest = "";
        public string bpjusti = "";
        public string swma = "";
        public string renew = "";
        public string totalprice = "";
        public string customerprice = "";
        public string totalright = "";
        public string totalright2 = "";
        public string customerright = "";
        public string customerright2 = "";
        public string sb_id_gestion = "";
        public string usuario = "";
        public string alerta_text = "";



        public void Main()//SB_Form_Main()
        {

            //descargar excel y CFG
            if (mail.GetAttachmentEmail("Special Bid", "Procesados", "Procesados Special Bid"))
            {
                Console.WriteLine(DateTime.Now + " > > > " + "Procesando...");
                Procesar_sb_form(roots.FilesDownloadPath + "\\" + roots.ExcelFile);
                Console.WriteLine(DateTime.Now + " > > > " + "Creando Estadisticas");
                
            }

        }

        public void Procesar_sb_form(string ruta)
        {

            int rows; string validacion; string respuesta = "";
            Console.WriteLine(DateTime.Now + " > > > " + "Abriendo Excel y Validando");
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlApp.Visible = false;

            xlWorkBook = xlApp.Workbooks.Open(ruta);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];

            rows = xlWorkSheet.UsedRange.Rows.Count;

            validacion = xlWorkSheet.Cells[2, 13].text.ToString();
            if (validacion.Substring(0, 1) != "x")
            {
                respuesta = "Utilizar la plantilla oficial de Special Bids";
                Console.WriteLine(DateTime.Now + " > > > " + "Devolviendo Solicitud");
                mail.SendHTMLMail(respuesta, new string[] { roots.BDUserCreatedBy }, roots.Subject, roots.CopyCC);
            }
            else
            {
                for (int i = 3; i <= rows; i++)
                {

                    projectname = xlWorkSheet.Cells[i, 2].text.ToString();

                    if (projectname != "")
                    {
                        projectname = projectname.Trim().ToUpper();
                        projectname = val.RemoveSpecialChars(projectname, 1);
                        #region extraer data
                        ibm_country = xlWorkSheet.Cells[i, 1].text.ToString().Trim().ToUpper();
                        useopp = xlWorkSheet.Cells[i, 3].text.ToString().Trim();
                        opp = xlWorkSheet.Cells[i, 4].text.ToString().Trim(); ;

                        useprevbid = xlWorkSheet.Cells[i, 5].text.ToString().Trim();
                        prevbid = xlWorkSheet.Cells[i, 6].text.ToString().Trim(); ;

                        priceupd = xlWorkSheet.Cells[i, 7].text.ToString().Trim();
                        justi = xlWorkSheet.Cells[i, 8].text.ToString().Trim(); ;

                        justi = val.RemoveSpecialChars(justi, 1);
                        justi = justi.Replace("\n", " ");
                        justi = justi.Replace("\r", " ");


                        customer = xlWorkSheet.Cells[i, 9].text.ToString().Trim().ToUpper();
                        brand = xlWorkSheet.Cells[i, 10].text.ToString().Trim().ToUpper();
                        justi2 = xlWorkSheet.Cells[i, 11].text.ToString().Trim().ToUpper();
                        justi2 = val.RemoveSpecialChars(justi2, 1);
                        justi2 = justi2.Replace("\r", " ");
                        justi2 = justi2.Replace("\n", " ");
                        addquest = xlWorkSheet.Cells[i, 12].text.ToString().Trim();

                        bpjusti = xlWorkSheet.Cells[i, 13].text.ToString().Trim().ToUpper();
                        bpjusti = val.RemoveSpecialChars(bpjusti, 1);
                        bpjusti = bpjusti.Replace("\n", " ");
                        bpjusti = bpjusti.Replace("\r", " ");
                        if (bpjusti.Length > 700) { bpjusti = bpjusti.Substring(0, 700); }

                        swma = xlWorkSheet.Cells[i, 14].text.ToString().Trim();
                        renew = xlWorkSheet.Cells[i, 15].text.ToString().Trim();
                        totalprice = xlWorkSheet.Cells[i, 16].text.ToString().Trim().ToUpper();
                        customerprice = xlWorkSheet.Cells[i, 17].text.ToString().Trim().ToUpper();

                        totalright = totalprice.Substring(totalprice.Length - 3, 3);
                        totalright2 = totalprice.Substring(totalprice.Length - 2, 2);
                        if (totalright.Substring(0, 1) == "." || totalright2.Substring(0, 1) == ".")
                        {
                            totalprice = totalprice.Replace(",", "");
                            totalprice = totalprice.Replace(".", ",");
                        }

                        customerright = customerprice.Substring(customerprice.Length - 3, 3);
                        customerright2 = customerprice.Substring(customerprice.Length - 2, 2);
                        if (customerright.Substring(0, 1) == "." || customerright2.Substring(0, 1) == ".")
                        {
                            customerprice = customerprice.Replace(",", "");
                            customerprice = customerprice.Replace(".", ",");
                        }

                        int punto;
                        punto = totalprice.IndexOf(".") + 1;
                        if (punto == 0) { totalprice += ".00"; }
                        punto = customerprice.IndexOf(".") + 1;
                        if (punto == 0) { customerprice += ".00"; }

                        #endregion

                        #region validar data

                        bool crear = val.ValidateSBForm(ibm_country, useopp, opp, useprevbid, prevbid, priceupd, justi, customer, brand, justi2,
                                      addquest, bpjusti, swma, renew, totalprice, customerprice);

                        if (crear == true)
                        {
                            try
                            {
                                Console.WriteLine(DateTime.Now + " > > > " + "Llenando formulario con Selenium");
                                sb.SbCreateFormWeb(ibm_country, projectname, useopp, opp, useprevbid, prevbid, priceupd, justi, customer, brand, justi2,
                                    addquest, bpjusti, swma, renew, totalprice, customerprice);
                                sb_number = roots.id_special_bid;
                                respuesta = respuesta + projectname + ": " + "Creado con Exito, " + sb_number + "<br>";
                                //enviar respuesta al usuario con el id de special bid

                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(DateTime.Now + " > > > " + "Error al crear el formulario");
                                respuesta = respuesta + projectname + ": " + "Error al crear el formulario" + "<br>";
                                Console.WriteLine(ex.ToString());
                                System.Threading.Thread.Sleep(1000);
                                proc.KillProcess("chromedriver", true);
                                proc.KillProcess("chrome", true);
                            }
                        }
                        else
                        {
                            respuesta = respuesta + projectname + ": " + "Error en el formulario, verifique la informacion" + "<br>";

                        }
                        #endregion
                        log.LogDeCambios("Creacion", roots.BDProcess, roots.BDUserCreatedBy, "Crear Special Bid Form", respuesta + ": " + sb_number, roots.Subject);
                    }


                } //for por cada fila del excel

                Console.WriteLine(DateTime.Now + " > > > " + "Respondiendo solicitud");
                if (respuesta.Contains("Error"))
                {
                    string[] cc = { "dmeza@gbm.net" };
                    mail.SendHTMLMail(respuesta, new string[] {"appmanagement@gbm.net"}, roots.Subject, cc);
                }
                else
                {

                    mail.SendHTMLMail(respuesta + "<br>" + "<br>" + "Haga Click <a href=\"https://extbasicbpmsprd.podc.sl.edst.ibm.com/bpms/\"" + ">aqui</a> para ver el documento", new string[] { roots.BDUserCreatedBy }, roots.Subject,  roots.CopyCC);
                }


            }
            xlApp.DisplayAlerts = false;
            xlApp.Workbooks.Close();
            xlApp.Quit();
            proc.KillProcess("EXCEL", true);
            proc.KillProcess("chromedriver", true);

        }




    }
}

