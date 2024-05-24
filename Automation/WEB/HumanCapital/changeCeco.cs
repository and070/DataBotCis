using DataBotV5.App.Global;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.ActiveDirectory;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Web;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataBotV5.Automation.WEB.HumanCapital
{
    public class changeCeco
    {
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        SapVariants sap = new SapVariants();
        MsExcel excel = new MsExcel();
        Rooting root = new Rooting();
        Stats stats = new Stats();
        CRUD crud = new CRUD();
        Log log = new Log();
        string mandante = "ERP"; 
        string mandanteSS = "PRD";
        string respFinal = "";

        /// <summary>
        /// 
        /// </summary>
        public void Main()
        {
            //utilizar en caso de que el robot utilice SAP Logon GUI
            if (!sap.CheckLogin(mandante))
            {
                DataTable excelDt = crud.Select( "SELECT NewCecoPosition.*, Positions.name FROM `NewCecoPosition` INNER JOIN Positions ON  NewCecoPosition.positionNumber = Positions.id WHERE statusBot = 1 LIMIT 1", "new_position_db");
                //Leer email y extraer adjunto
                if (excelDt.Rows.Count > 0)
                {
                    console.WriteLine("Processing...");
                    //utilizar en caso de que el robot utilice SAP Logon GUI
                    sap.BlockUser(mandante, 1);
                    ///Procesar------------------
                    Process(excelDt);
                    ///--------------------------
                    //utilizar en caso de que el robot utilice SAP Logon GUI
                    sap.BlockUser(mandante, 0);


                    //crear estadisticas
                    console.WriteLine("Creando estadísticas...");
                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }

                }
            }
        }
        /// <summary>
        ///
        /// </summary>
        /// <param name="ExcelFile">el excel que envía el usuario por email outlook</param>
        private void Process(DataTable ExcelFile)
        {
            #region private variables
            //PLantilla en html para el envío de email
            string htmlEmail = Properties.Resources.emailtemplate1;
            //variable titulo del cuerpo del correo
            string htmlSubject = "Resultados";
            //variable contenido del correo: texto, cuadros, tablas, imagenes, etc
            string htmlContents = "";


            #endregion

            #region loop each excel row
            console.WriteLine("Foreach Excel row...");
            string respuesta = "";
            string mensaje_devolucion = "";
            string response_failure = "";
            string userCreatedBy = "";
            string positionName = "";
            bool validar_lineas = true;
            foreach (DataRow rRow in ExcelFile.Rows)
            {

                string ID = rRow["ID"].ToString().Trim();
                string userID = "";
                try
                {
                    DataTable files = crud.Select( $@"SELECT NewCecoPositionFiles.*, 
                                                                    Files.name,
                                                                    Files.file as fileBlop
                                                                    FROM NewCecoPositionFiles
                                                                    INNER JOIN Files ON NewCecoPositionFiles.files = Files.id
                                                                    WHERE NewCecoPositionFiles.newCecoPosition = {ID} 
                                                                    GROUP BY NewCecoPositionFiles.ID;",
                                                            "new_position_db");
                    #region robot Process
                    //subir workflow
                    #region extraer data y crear comentario

                    string comments = rRow["comments"].ToString().Trim();
                    userID = rRow["userID"].ToString().Trim();
                    string changeRequestDate = DateTime.Parse(rRow["changeRequestDate"].ToString().Trim()).ToString("dd/MM/yyyy");
                    userCreatedBy = rRow["createdBy"].ToString().Trim();
                    string nCeco = rRow["cecoN"].ToString().Trim();
                    positionName = rRow["name"].ToString().Trim();

                    comments = comments.Replace("\n", "");
                    comments = comments.Replace("\r", "");

                    string wfComments = "Nuevo CeCo:" + "\r\n";
                    wfComments = wfComments + nCeco + "\r\n";
                    wfComments = wfComments + "\r\n";
                    wfComments = wfComments + "Fecha de vigencia:" + "\r\n";
                    wfComments = wfComments + changeRequestDate + "\r\n";
                    wfComments = wfComments + "\r\n";
                    wfComments = wfComments + "Comentarios de la solicitud:" + "\r\n";
                    wfComments = wfComments + comments + "\r\n";
                    wfComments = wfComments + "\r\n";
                    wfComments = wfComments + "Usuario Solicitante: " + "\r\n";
                    wfComments = wfComments + userCreatedBy;

                    #endregion


                    #region cargar la posicion en SAP
                    console.WriteLine(" Cargar la posicion en SAP");

                    sap.LogSAP(mandante.ToString());

                    // SAP_Variants.frame.Iconify();
                    ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nZHR_WF19";
                    SapVariants.frame.SendVKey(0);
                    ((SAPFEWSELib.GuiComboBox)SapVariants.session.FindById("wnd[0]/usr/cmbZHRCP019-TPO_SOL")).Key = "13";
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtZHRCP019-EMPLEADO")).Text = userID;
                    SapVariants.frame.SendVKey(0);
                    ((SAPFEWSELib.GuiTextedit)SapVariants.session.FindById("wnd[0]/usr/subSUB_SCREEN:ZHRPG_WF_OTROS:1001/cntlCTRL_TEXT/shellcont/shell")).Text = wfComments;
                    #region Subir archivos a SAP
                    if (files.Rows.Count > 0)
                    {
                        foreach (DataRow fRow in files.Rows)
                        {
                            //convertir blop
                            string path = root.FilesDownloadPath + "\\" + fRow["name"].ToString();
                            if (File.Exists(path))
                            {
                                File.Delete(path);
                            }
                            byte[] binDate = (byte[])fRow["fileBlop"];
                            File.WriteAllBytes(path, binDate);


                            //adjuntar en sap
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/subSUB_SCREEN:ZHRPG_WF_OTROS:1001/ctxtFILE")).Text = "";
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/subSUB_SCREEN:ZHRPG_WF_OTROS:1001/ctxtFILE")).SetFocus();
                            SapVariants.frame.SendVKey(4);
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtDY_PATH")).Text = root.FilesDownloadPath + "\\";
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtDY_FILENAME")).Text = fRow["name"].ToString();
                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                            try
                            {
                                //EN CASO DE QUE SALGA UN POP DE REPETIDO
                                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                            }
                            catch (Exception)
                            {

                            }
                        }
                    }
                    #endregion
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[20]")).Press(); //ENVIAR WORKFLOW
                    try
                    {
                        respuesta = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString();
                    }
                    catch (Exception)
                    { }

                    crud.Update($"UPDATE NewCecoPosition SET statusBot = 0 WHERE ID = {ID}", "new_position_db");

                    string respo = $"Se actualiza el CECO de la posición {positionName} en SAP y S&S.";
                    log.LogDeCambios("Actualización", root.BDProcess,  userCreatedBy, "Actualizar CECO", respo, "");
                    respFinal = respFinal + "\\n" + respo;


                }
                catch (Exception ex)
                {
                    try
                    { mensaje_devolucion = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString(); }
                    catch (Exception) { }
                    response_failure = new ValidateData().LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, 0);
                    console.WriteLine(" Finishing process " + response_failure);
                    respuesta = userID + ": " + mensaje_devolucion + "<br>" + "<br>" + ex.ToString();
                    response_failure = ex.ToString();
                    validar_lineas = false;
                    crud.Update($"UPDATE NewCecoPosition SET statusBot = 2 WHERE ID = {ID}", "new_position_db");
                    continue;
                }
                sap.KillSAP();

                #endregion




                //log de cambios
                //log.LogDeCambios("Creacion", root.BDProcess, root.Solicitante, logText, "", "");
                #endregion
            }
            #endregion


            #region SendEmail


            if (respuesta == "")
            {
                respuesta = "Hubo un error al crear el workflow de cambio de CeCo, por favor verifique la data";
                validar_lineas = false;
            }

            root.Subject = $"Solicitud para actualización de CeCo en posición: {positionName}";

            if (validar_lineas == false)
            {
                //enviar email de repuesta de error
                //string[] cc = { userCreatedBy, "gvillalobos@gbm.net" };
                string[] cc = { "kcarvajal@gbm.net", "dmeza@gbm.net" , userCreatedBy };
                //enviar email de repuesta de exito
                htmlSubject = "Error al crear Flujo de Aprobación de Centro de Costos Creado en SAP";
                htmlContents = "";
                htmlEmail = htmlEmail.Replace("{subject}", htmlSubject).Replace("{cuerpo}", respuesta).Replace("{contenido}", htmlContents);
                console.WriteLine("Send Email...");
                mail.SendHTMLMail(htmlEmail, new string[] {"appmanagement@gbm.net"}, "Error: " + root.Subject, cc, null);
                root.BDUserCreatedBy = userCreatedBy;

            }
            else
            {

                //enviar email de repuesta de exito
                htmlSubject = "Nuevo Flujo de Aprobación de Centro de Costos Creado en SAP";
                htmlContents = "";
                htmlEmail = htmlEmail.Replace("{subject}", htmlSubject).Replace("{cuerpo}", respuesta).Replace("{contenido}", htmlContents);
                console.WriteLine("Send Email...");
                mail.SendHTMLMail(htmlEmail, new string[] { userCreatedBy }, root.Subject, null, null);
                root.BDUserCreatedBy = userCreatedBy;

            }

            #endregion

            root.requestDetails = respFinal;


        }




    }
}
