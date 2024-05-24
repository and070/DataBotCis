using System;
using System.IO;
using SAP.Middleware.Connector;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Projects.BusinessSystem;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using System.Collections.Generic;

namespace DataBotV5.Automation.RPA.IBMPoe
{
    /// <summary>
    /// Clase RPA Automation encargada de la facturación de IBM Poe.
    /// </summary>
    class IBMPoeBilling
    {
        Stats estadisticas = new Stats();
        Credentials cred = new Credentials();
        ConsoleFormat console = new ConsoleFormat();
        Rooting root = new Rooting();
        SapVariants sap = new SapVariants();
        ProcessInteraction proc = new ProcessInteraction();
        BsSQL bsql = new BsSQL();
        MailInteraction mail = new MailInteraction();
        Log log = new Log();
        string respuesta = "";
        string respuesta2 = "";
        string validacion = "";
        string purch_order = "";
        string pdf_fileA = ""; string pdf_fileB = "";
        string salesorder;
        string mandante = "ERP";
        string respFinal = "";


        public void Main()
        {
            //revisa si el usuario RPAUSER esta abierto
            pdf_fileA = ""; pdf_fileB = "";
            if (mail.GetAttachmentEmail("Solicitudes POE", "Procesados", "Procesados POE"))
            {
                console.WriteLine("Procesando...");
                if (!Directory.Exists(root.FilesDownloadPath + "\\" + "POE (No eliminar)"))
                {
                    Directory.CreateDirectory(root.FilesDownloadPath + "\\" + "POE (No eliminar)");
                }
                #region eliminar y copiar archivo
                if (File.Exists(root.FilesDownloadPath + "\\" + "POE (No eliminar)" + "\\" + root.ExcelFile))
                {
                    File.Delete(root.FilesDownloadPath + "\\" + "POE (No eliminar)" + "\\" + root.ExcelFile);
                }
                pdf_fileA = root.FilesDownloadPath + "\\" + root.ExcelFile; //POE (No eliminar)
                pdf_fileB = root.FilesDownloadPath + "\\" + "POE (No eliminar)" + "\\" + root.ExcelFile;
                File.Copy(pdf_fileA, pdf_fileB);
                #endregion
                ProcessIBMPoe(root.FilesDownloadPath + "\\" + root.ExcelFile);

                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
            //}
        }

        public void ProcessIBMPoe(string route)
        {
            if (root.Subject.IndexOf(", 77") + 1 != 0)
            {
                respuesta = "no aplica";
                console.WriteLine("POE de Software, no aplica");

            }
            else
            {
                if (root.Subject.IndexOf("PoE-P.O.") + 1 != 0) //Correo B
                {
                    console.WriteLine("Ingresando Correo B");
                    try
                    {
                        purch_order = root.Subject.Substring(root.Subject.IndexOf("PoE-P.O.") + 9, 10);
                    }
                    catch (Exception)
                    {
                        purch_order = root.Subject.Substring(root.Subject.IndexOf("PoE-P.O.") + 9, 9);
                    }

                    //metodo para extraer el nombre del PDF del correo A para enviarlos x email con el PDF de este email
                    bsql.SqlIBMPoe("emailBPoe", purch_order, root.ExcelFile);

                    if (root.ibm_pdf2 == "no file")
                    {
                        //no encontro el file del correo A
                        console.WriteLine("no encontro el file del correo A");
                        string[] adjunto = { root.FilesDownloadPath + "\\" + root.ExcelFile };
                        mail.SendHTMLMail("PO: " + purch_order + "<br>" + "SO: " + salesorder + "<br> no encontro el file del correo A",
                            new string[] { "dmeza@gbm.net" },
                            "Error: " + root.Subject + " " + salesorder,
                            new string[] { "epiedra@gbm.net" },
                            adjunto);

                    }
                    else
                    {
                        //agregar al array el nombre del archivo del correo actual y el nombre del archivo del correo A que esta en la variable ibm_pdf2
                        salesorder = "";

                        #region buscar salesorder en SAP

                        try
                        {
                            console.WriteLine("Extraer SO de SAP");

                            Dictionary<string, string> parameters = new Dictionary<string, string>();
                            parameters["TIPO_DOC"] = "PO";
                            parameters["ID_DOC"] = purch_order;

                            IRfcFunction func1 = sap.ExecuteRFC(mandante, "ZFI_READ_PO", parameters);




                            if (func1.GetValue("RESPUESTA").ToString().Trim() == "NA")
                            {
                                salesorder = "";
                            }
                            else
                            {
                                salesorder = func1.GetValue("SO_PAIS").ToString().Trim();
                            }

                        }
                        catch (Exception)
                        { }
                        #endregion

                        //extraer el email de los usuarios y sus copias
                        console.WriteLine("Enviando Email a Facturacion");
                        string[] cc = bsql.EmailAddress(1);
                        //string[] cc = { root.f_copy1, root.f_copy2, root.f_copy3, root.f_copy4, root.f_copy5, root.f_copy6, root.f_copy7, root.f_copy8 };

                        string[] adjunto = { root.FilesDownloadPath + "\\" + root.ExcelFile, root.FilesDownloadPath + "\\" + "POE (No eliminar)" + "\\" + root.ibm_pdf2 };
                        mail.SendHTMLMail("PO: " + purch_order + "<br>" + "SO: " + salesorder, new string[] { root.f_sender }, root.Subject + " " + salesorder, cc, adjunto);

                        //log de cambios
                        log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Enviar email de POE", purch_order, root.Subject);
                        respFinal = respFinal + "\\n" + purch_order;


                    }
                }
                else if (root.Subject.IndexOf("Miami Direct, Inc.") + 1 != 0) //correo A
                {
                    console.WriteLine("Ingresando Correo A");
                    purch_order = root.Subject.Substring(10, 10);
                    //metodo para guardar en base de datos el nombre del PDF del correo A y su PO
                    bsql.SqlIBMPoe("emailAPoe", purch_order, root.ExcelFile);
                }
            }

            root.requestDetails = respFinal;

        }


    }
}
