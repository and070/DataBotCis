using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using SAP.Middleware.Connector;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Root;
using System.Data;
using System;

namespace DataBotV5.Automation.RPA.Materials
{
    /// <summary>
    /// Clase RPA Automation encargada de cancelar una solicitud de material.
    /// </summary>
    class CancelMR
    {
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        ValidateData val = new ValidateData();
        SapVariants sap = new SapVariants();
        MsExcel excel = new MsExcel();
        Rooting root = new Rooting();
        Stats stats = new Stats();
        Log log = new Log();
        string respFinal = "";


        string mand = "ERP";

        public void Main()
        {
            if (mail.GetAttachmentEmail("Solicitudes cancelar MR", "Procesados", "Procesados cancelar MR"))
            {
                DataTable excelDt = excel.GetExcel(root.FilesDownloadPath + "\\" + root.ExcelFile);
                CancelMRProcessing(excelDt);
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }

        public void CancelMRProcessing(DataTable excelDt)
        {
            string responseFailure = "";

            DataTable dtResponse = new DataTable();
            dtResponse.Columns.Add("SD Doc.");
            dtResponse.Columns.Add("Item");
            dtResponse.Columns.Add("Material");
            dtResponse.Columns.Add("Respuesta");

            string validation = excelDt.Columns[0].ColumnName;

            RfcDestination destination = sap.GetDestRFC(mand);
            IRfcFunction func = destination.Repository.CreateFunction("ZFI_CANCEL_MR");
            IRfcTable matRequest = func.GetTable("IN");
            IRfcTable mRequestOut = func.GetTable("OUT");

            try
            {
                if (validation.Substring(0, 1) == "x") //pass excel = cancelarmr
                {
                    foreach (DataRow row in excelDt.Rows)
                    {
                        string sdDoc = row[0].ToString().Trim();
                        string item = row[1].ToString().Trim();
                        string material = row[2].ToString().Trim();
                        string qty = row[3].ToString().Trim();

                        if (sdDoc != "")
                        {
                            if (material != "")
                            {
                                matRequest.Append();
                                matRequest.SetValue("VBELN_VA", sdDoc);
                                matRequest.SetValue("MATNR", material.ToUpper());
                                matRequest.SetValue("POSNR_VA", item);
                            }
                        }
                    }

                    #region SAP
                    console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                    func.Invoke(destination);

                    for (int i = 0; i < mRequestOut.RowCount; i++)
                    {
                        //construir el excel
                        DataRow drResponse = dtResponse.NewRow();

                        drResponse["SD Doc."] = mRequestOut[i].GetValue("VBELN_VA").ToString();
                        drResponse["Item"] = mRequestOut[i].GetValue("POSNR_VA").ToString();
                        drResponse["Material"] = mRequestOut[i].GetValue("MATNR").ToString();

                        //Arreglar mensaje
                        string res = mRequestOut[i].GetValue("MESSAGE").ToString().Trim();
                        res = res.Replace("ABGRU", "Reason for Rejection");
                        res = res.Replace("MATNR", "Material");
                        res = res.Replace("VBAPKOM", "Item");

                        drResponse["Respuesta"] = res;

                        dtResponse.Rows.Add(drResponse);

                        //log de base de datos
                        log.LogDeCambios("Modificacion", root.BDProcess, root.BDUserCreatedBy, "Cancelar Material Request", mRequestOut[i].GetValue("MESSAGE").ToString().Trim(), root.Subject);
                        respFinal = respFinal + "\\n" + "Cancelar Material Request : " + mRequestOut[i].GetValue("MESSAGE").ToString().Trim();

                    }
                    #endregion

                }
                else
                    mail.SendHTMLMail("Utilizar la plantilla oficial", new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);
            }
            catch (Exception ex)
            {
                responseFailure = ex.Message;
                console.WriteLine("Finishing process " + responseFailure);
            }

            console.WriteLine("Respondiendo solicitud");

            if (responseFailure != "")//enviar email de repuesta de error
                mail.SendHTMLMail(responseFailure, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject);
            else//enviar email de repuesta de éxito
                mail.SendHTMLMail(val.ConvertDataTableToHTML(dtResponse), new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);


            root.requestDetails = respFinal;

        }
    }
}
