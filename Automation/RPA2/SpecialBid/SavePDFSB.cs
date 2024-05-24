using System;
using System.IO;
using System.Text.RegularExpressions;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Parsing;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Projects.BusinessSystem;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Projects.SpecialBidPDF;
using DataBotV5.App.Global;
using OpenQA.Selenium;
using DataBotV5.Logical.Encode;
using DataBotV5.Data.Database;
using System.Data;
using DataBotV5.Logical.MicrosoftTools;

namespace DataBotV5.Automation.RPA2.SpecialBid
{
    /// <summary>
    /// Clase RPA Automation encargada de guardar en PDF una licitación especial.
    /// </summary>
    class SavePDFSB
    {
        Rooting roots = new Rooting();
        ValidateData val = new ValidateData();
        SBPDFSave sb = new SBPDFSave();
        ConsoleFormat console = new ConsoleFormat();
        ProcessInteraction proc = new ProcessInteraction();
        MailInteraction mail = new MailInteraction();
        Rooting root = new Rooting();
        BsSQL bsql = new BsSQL();
        Log log = new Log();
        CRUD crud = new CRUD();
        Stats estadisticas = new Stats();
        MsExcel ex = new MsExcel();
        string respFinal = "";


        public void Main()
        {
            //extrae el body del email
            if (mail.GetAttachmentEmail("Solicitudes Save SB", "Procesados", "Procesados Save SB"))
            {
                console.WriteLine("Procesando...");
                ProcessSBPdf();

                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }

        }

        public void ProcessSBPdf()
        {

            string pdf_name_save = ""; string folder_name = ""; string pdf_fileA = ""; string pdf_fileB = ""; string href = ""; string folder_mes = "";
            string folder_ano = ""; string folder_pais = ""; string cliente = ""; int num_client = 0; string ruta = ""; string ruta1 = ""; string ruta2 = "";
            string html_tag = ""; string fldrpath; string body; string respuesta = ""; string sb_save_num = ""; string[] folder_split; string[] link;
            string[] cc = bsql.EmailAddress(3);
            //string[] cc = { roots.f_copy1, roots.f_copy2, roots.f_copy3, roots.f_sender };
            body = roots.Email_Body;
            string sub = root.Subject;
            string solicitante = root.BDUserCreatedBy;
            if (body.Contains("#de Special Bid:"))
            {
                console.WriteLine("Descargando el PDF con Selenium");
                #region extraer el número del SB

                Regex reg;
                reg = new Regex("[*'\"_&+^><@]");
                html_tag = reg.Replace(body, string.Empty);

                string[] stringSeparators0 = new string[] { "#de Special Bid:" };
                link = html_tag.Split(stringSeparators0, StringSplitOptions.None);
                sb_save_num = link[1].ToString().Trim();
                if (sb_save_num.Length > 10)
                {
                    sb_save_num = sb_save_num.Substring(0, 10).Trim();
                }
                if (sb_save_num != string.Empty)
                {
                    if (sb_save_num.Length == 8 || sb_save_num.Substring(0, 2) != "00")
                    {
                        sb_save_num = "00" + sb_save_num.Substring(0, 8);
                    }
                }

                #endregion



                #region primeras carpetas
                ruta1 = @"\\Fs01\bs\SB";

                if (!Directory.Exists(ruta1))
                {
                    Directory.CreateDirectory(ruta1);
                }

                ruta2 = @"\\Fs01\bs\SB\SB IBM SW";

                if (!Directory.Exists(ruta2))
                {
                    Directory.CreateDirectory(ruta2);
                }
                #endregion

                ruta = @"\\Fs01\bs\SB\SB IBM SW";

                if (sb_save_num == "")
                {
                    respuesta = "No se encontro el # del SB";
                }
                else
                {
                    try
                    {
                        #region selenium
                        string pdf = "";
                        try
                        {
                            pdf = sb.SavePdfSb(sb_save_num,  solicitante , sub);
                        }
                        catch (Exception ex)
                        {
                            proc.KillProcess("chromedriver", true);
                            proc.KillProcess("chrome", true);
                            pdf = sb.SavePdfSb(sb_save_num,  solicitante , sub);
                        }
                        #endregion
                        if (pdf == "error")
                        {
                            string imagen64 = "";
                            byte[] errorImage = System.IO.File.ReadAllBytes(roots.FilesDownloadPath + "\\errorDownloadSpecialBidPdf.png");
                            using (BinaryFiles bf = new BinaryFiles())
                            {
                                imagen64 = bf.BinaryHTMLImage(errorImage);
                            }

                            string msj = $"No se pudo descargar el PDF del SpecialBid: {sb_save_num} debido a: </br>{imagen64}";
                            string html = Properties.Resources.emailtemplate1;
                            html = html.Replace("{subject}", $"Error al guardar documento pdf relacionado al Special Bid: {sb_save_num}");
                            html = html.Replace("{cuerpo}", msj);
                            html = html.Replace("{contenido}", "");
                            mail.SendHTMLMail(html, new string[] { solicitante }, sub, cc, null);
                            return;
                        }
                        else if (pdf == "No existe PDF")
                        {
                            string msj = $"No se encontró un documento asociado al SpecialBid: {sb_save_num}";
                            string html = Properties.Resources.emailtemplate1;
                            html = html.Replace("{subject}", $"Guardar documento pdf relacionado al Special Bid: {sb_save_num}");
                            html = html.Replace("{cuerpo}", msj);
                            html = html.Replace("{contenido}", "");
                            mail.SendHTMLMail(html, new string[] { solicitante }, sub, cc, null);
                            return;
                        }
                        pdf_name_save = "Channel Bid Notification PDF " + sb_save_num + ".pdf";
                        console.WriteLine("Creando la carpeta");
                        #region archivo_existe
                        //saber si el archivo existe con el nombre anterior o no.
                        if (File.Exists(roots.FilesDownloadPath + "\\" + pdf_name_save))
                        {
                            pdf_name_save = "Channel Bid Notification PDF " + sb_save_num + ".pdf";
                            respuesta = "Archivo descargado";
                        }
                        else
                        {
                            pdf_name_save = "customer quote PDF " + sb_save_num + ".pdf";
                            if (File.Exists(roots.FilesDownloadPath + "\\" + pdf_name_save))
                            {
                                pdf_name_save = "customer quote PDF " + sb_save_num + ".pdf";
                                respuesta = "Archivo descargado";
                            }
                            else
                            {
                                respuesta = "error al descargar el PDF";
                            }
                        }

                        #endregion

                        if (respuesta != "error al descargar el PDF")
                        {
                            #region folder año
                            //busca si la carpeta del año existe
                            folder_ano = DateTime.Now.Year.ToString();
                            fldrpath = ruta + "\\" + folder_ano;
                            if (!Directory.Exists(fldrpath))
                            {
                                Directory.CreateDirectory(fldrpath);
                            }
                            #endregion

                            #region folder mes
                            //"busca si la carpeta del mes existe"

                            switch (DateTime.Now.Month)
                            {
                                case 1:
                                    folder_mes = "Enero";
                                    break;
                                case 2:
                                    folder_mes = "Febrero";
                                    break;
                                case 3:
                                    folder_mes = "Marzo";
                                    break;
                                case 4:
                                    folder_mes = "Abril";
                                    break;
                                case 5:
                                    folder_mes = "Mayo";
                                    break;
                                case 6:
                                    folder_mes = "Junio";
                                    break;
                                case 7:
                                    folder_mes = "Julio";
                                    break;
                                case 8:
                                    folder_mes = "Agosto";
                                    break;
                                case 9:
                                    folder_mes = "Setiembre";
                                    break;
                                case 10:
                                    folder_mes = "Octubre";
                                    break;
                                case 11:
                                    folder_mes = "Noviembre";
                                    break;
                                case 12:
                                    folder_mes = "Diciembre";
                                    break;
                                default:
                                    folder_mes = "Enero";
                                    break;
                            }

                            fldrpath = fldrpath + "\\" + folder_mes;
                            if (!Directory.Exists(fldrpath))
                            {
                                Directory.CreateDirectory(fldrpath);
                            }
                            #endregion

                            #region Folder pais
                            PdfLoadedDocument loadedDocument = new PdfLoadedDocument(roots.FilesDownloadPath + "\\" + pdf_name_save);
                            PdfPageBase pages = loadedDocument.Pages[0];
                            string extractedTexts = pages.ExtractText();
                            loadedDocument.Close();

                            int ipais;
                            ipais = extractedTexts.IndexOf("Country:");
                            if (ipais > 0)
                            {

                                string[] stringSeparators = new string[] { "Country:" };
                                folder_split = extractedTexts.Split(stringSeparators, StringSplitOptions.None);
                                folder_pais = folder_split[1].ToString();
                                folder_pais = folder_pais.Replace("\n", " ").Trim();
                                folder_pais = folder_pais.Replace("\r", " ");
                                if (folder_pais != "")
                                {
                                    folder_pais = folder_pais.Substring(0, 2);
                                }

                                switch (folder_pais)
                                {
                                    case "CR":
                                        folder_pais = "Costa Rica";
                                        break;
                                    case "DO":
                                        folder_pais = "Republica Dominicana";
                                        break;
                                    case "GT":
                                        folder_pais = "Guatemala";
                                        break;
                                    case "NI":
                                        folder_pais = "Nicaragua";
                                        break;
                                    case "PA":
                                        folder_pais = "Panama";
                                        break;
                                    case "US":
                                        folder_pais = "USA";
                                        break;
                                    case "SV":
                                        folder_pais = "El Salvador";
                                        break;
                                    case "VE":
                                        folder_pais = "Venezuela";
                                        break;
                                    case "CO":
                                        folder_pais = "Colombia";
                                        break;
                                    case "BZ":
                                        folder_pais = "Honduras";
                                        break;
                                    case "HN":
                                        folder_pais = "Honduras";
                                        break;
                                    default:
                                        folder_pais = "Otro";
                                        break;
                                }

                                fldrpath = fldrpath + "\\" + folder_pais;
                                if (!Directory.Exists(fldrpath))
                                {
                                    Directory.CreateDirectory(fldrpath);
                                }
                            }
                            else
                            {
                                fldrpath = fldrpath + "\\" + "Otro";
                                if (!Directory.Exists(fldrpath))
                                {
                                    Directory.CreateDirectory(fldrpath);
                                }
                            }
                            #endregion

                            #region folder cliente
                            //crea la carpeta en Z con el nombre del SB y el cliente

                            PdfLoadedDocument loadedDocument2 = new PdfLoadedDocument(roots.FilesDownloadPath + "\\" + pdf_name_save);
                            PdfPageBase pages2 = loadedDocument2.Pages[0];
                            string extractedTexts2 = pages2.ExtractText();
                            loadedDocument.Close();
                            int ipais2;
                            ipais2 = extractedTexts.IndexOf("Customer Name:");
                            if (ipais2 > 0)
                            {

                                string[] stringSeparators2 = new string[] { "Customer Name:" };
                                folder_split = extractedTexts.Split(stringSeparators2, StringSplitOptions.None);
                                num_client = folder_split[1].ToString().IndexOf("Address:") - 1;

                                if (folder_split[1] != "")
                                {
                                    cliente = folder_split[1].ToString().Substring(0, num_client);
                                }

                                cliente = cliente.Replace("\n", " ").Trim();
                                cliente = cliente.Replace("\r", " ").Trim();
                                folder_name = "SB" + sb_save_num + " - " + cliente;
                                Regex reg2;
                                reg2 = new Regex("[*:'\"_&+^><@]");
                                folder_name = reg2.Replace(folder_name, string.Empty);

                                fldrpath = fldrpath + "\\" + folder_name;
                                if (!Directory.Exists(fldrpath))
                                {
                                    Directory.CreateDirectory(fldrpath);
                                }
                            }
                            else
                            {
                                fldrpath = fldrpath + "\\" + "Otro";
                                if (!Directory.Exists(fldrpath))
                                {
                                    Directory.CreateDirectory(fldrpath);
                                }
                            }
                            #endregion

                            #region elimina y copia

                            //elimina el archivo que ya este dento de la carpeta

                            if (File.Exists(fldrpath + "\\" + pdf_name_save))
                            {
                                File.Delete(fldrpath + "\\" + pdf_name_save);
                            }

                            // mueve el archivo

                            try
                            {
                                console.WriteLine("Copiando PDF y enviando respuesta");
                                pdf_fileA = roots.FilesDownloadPath + "\\" + pdf_name_save;
                                pdf_fileB = fldrpath + "\\" + pdf_name_save;
                                File.Copy(pdf_fileA, pdf_fileB);

                                string msj = "Se ha descargado el PDF, y se guardó en el Z, ruta: " + "<br>" + fldrpath;
                                string html = Properties.Resources.emailtemplate1;
                                html = html.Replace("{subject}", $"Guardar documento pdf relacionado al Special Bid: {sb_save_num}");
                                html = html.Replace("{cuerpo}", msj);
                                html = html.Replace("{contenido}", "");
                                mail.SendHTMLMail(html, new string[] { solicitante }, sub, cc, null);
                                log.LogDeCambios("Creacion", roots.BDProcess, solicitante, "Guardar PDF Special Bid", "Se ha descargado el PDF, y se guardo en el Z del specialBid:" + sb_save_num, sub);
                                respFinal = respFinal + "\\n" + "Se guarda-descarga pdf relacionado al Special Bid: " + sb_save_num;

                            }
                            catch (Exception ex)
                            {
                                console.WriteLine("Error: " + ex.ToString());
                                //enviar email de error al copiar el archivo
                                mail.SendHTMLMail(respuesta + "<br>" + "error al copiar el archivo a Z: " + ex.ToString(), new string[] {"appmanagement@gbm.net"}, sub, cc);
                            }


                            #endregion


                        }
                        else
                        {
                            console.WriteLine("Error al descargar el PDF");
                            respuesta = "Error al descargar el PDF";
                            //enviar email de error a datos maestros
                            mail.SendHTMLMail(respuesta, new string[] {"appmanagement@gbm.net"}, sub, cc);

                        }


                    }
                    catch (Exception ex)
                    {
                        console.WriteLine("Error: " + ex.ToString());
                        respuesta = "Error al descargar el PDF";
                        proc.KillProcess("chromedriver", true);
                        proc.KillProcess("chrome", true);
                        mail.SendHTMLMail(respuesta + " - " + ex.ToString(), new string[] {"appmanagement@gbm.net"}, sub, cc);
                    }





                }

            }

            root.requestDetails = respFinal;

        }


    }
}
