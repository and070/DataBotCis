using DataBotV5.Logical.Projects.UserSAM;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Data;
using System.Linq;
using System;

namespace DataBotV5.Automation.ICS.SAM

{
    /// <summary>
    /// Clase ICS Automation encargada de la gestión de usuarios en SAM.
    /// </summary>
    class SAMUsers
    {
        readonly FusionAuthInteraction fai = new FusionAuthInteraction();
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly ValidateData val = new ValidateData();
        readonly MsExcel excel = new MsExcel();
        readonly Rooting root = new Rooting();
        readonly Log log = new Log();

        public void Main()
        {
            console.WriteLine("Descargando archivo");

            if (mail.GetAttachmentEmail("Usuarios SAM", "Procesados", "Procesados Usuarios SAM"))
            {
                console.WriteLine("Procesando...");
                DataTable excelDt = excel.GetExcel(root.FilesDownloadPath + "\\" + root.ExcelFile);
                ProcessSAM(excelDt);

                console.WriteLine("Creando Estadísticas");
                using (Stats stats = new Stats()) { stats.CreateStat(); }
            }
        }

        public void ProcessSAM(DataTable excelDt)
        {
            string respFinal = "";
            string customerCol = "CUST ID";
            string firstNameCol = "NOMBRE DEL CONTACTO";
            string lastNameCol = "APELLIDO DEL CONTACTO";
            string emailCol = "xCORREO ELECTRONICO / USUARIO SAM";
            string passCol = "Contraseña";
            bool sendIcs = false;

            string excelValidation = excelDt.Columns[emailCol].ColumnName;
            excelDt.Columns.Add("Resultado");

            if (String.Concat(excelValidation.Take(1)) == "x")//Plantilla correcta, continúe las validaciones
            {
                foreach (DataRow excelRow in excelDt.Rows)
                {
                    string rowRes = "";
                    string email = excelRow[emailCol].ToString().Trim();

                    if (email != "")
                    {
                        string customer = excelRow[customerCol].ToString().Trim();

                        if (customer != "")
                        {
                            string firstName = excelRow[firstNameCol].ToString().Trim().ToUpper();
                            string lastName = excelRow[lastNameCol].ToString().Trim().ToUpper();
                            string passReq = excelRow[passCol].ToString().Trim();

                            lastName = val.RemoveSpecialChars(lastName, 1);
                            firstName = val.RemoveSpecialChars(firstName, 1);
                            string fullName = firstName + " " + lastName;

                            int index = lastName.IndexOf(" ");
                            string passLName;
                            if (index > -1)
                                passLName = lastName.Substring(0, index);
                            else
                                passLName = lastName;

                            string generatedPass = firstName.Substring(0, 1) + passLName.ToLower() + DateTime.Now.Year;

                            #region SAP

                            string[] sapRes = CreateUserInSap(email, customer, firstName, lastName);
                            string fmRes = sapRes[0];
                            string fmSapContactId = sapRes[1];
                            string fmMsgResponse = sapRes[2];

                            try
                            {
                                #region Procesar Salidas del FM

                                if (fmRes.ToLower().Contains("contacto actualizado") || fmMsgResponse.ToLower().Contains("business partner was created with number"))
                                {
                                    console.WriteLine("Creando Usuario en Fusion Auth");

                                    FaRegistrationData registration = new FaRegistrationData
                                    {
                                        UserName = email,
                                        ApplicationId = "f6676728-74b5-4364-80a4-bc42b85d5879", //el id de External SAM
                                        Roles = new List<string>
                                        {
                                            "Tiquetes",
                                            "Productos",
                                            "General",
                                            "Finanzas",
                                            "Contratos",
                                            "Admin"
                                        }
                                    };

                                    FaUserData userData = new FaUserData
                                    {
                                        Email = email,
                                        FirstName = firstName,
                                        FullName = fullName,
                                        LastName = lastName,
                                        Password = generatedPass,
                                        UserName = email,
                                        Registration = registration
                                    };

                                    string faRes = fai.CreateUser(userData);

                                    if (faRes == "OK")
                                    {
                                        rowRes = generatedPass;
                                        log.LogDeCambios("", "",  root.BDUserCreatedBy , "Crear Usuarios SAM", customer + ": " + email + " - " + generatedPass, faRes);
                                        respFinal += "\\n" + "Crear Usuario SAM " + customer + ": " + email + " - " + generatedPass;

                                    }
                                    else if (faRes.Contains("A User with username [") && faRes.Contains("] already exists."))
                                    {
                                        rowRes = email + "(" + fmSapContactId + ") ya existe";
                                        log.LogDeCambios("", "", root.BDUserCreatedBy, "Crear Usuarios SAM", rowRes, faRes);
                                        respFinal += "\\n" + "Crear Usuario SAM " + customer + ": " + email + " - " + generatedPass + ": " + faRes;
                                    }
                                    else
                                    {
                                        sendIcs = true;
                                        rowRes = "<b>Error, no se creo el Usuario en Fusion Auth: </b><br>" + faRes + "";
                                    }
                                }
                                else if (fmRes.ToLower().Contains("se agrego el cliente"))
                                {
                                    if (passReq.ToLower() == "x")
                                    {
                                        console.WriteLine("Cambiar pass en Fusion Auth");

                                        FaUserData userData = new FaUserData
                                        {
                                            UserName = email,
                                            Password = generatedPass,
                                        };

                                        userData.Id = fai.GetUserId(userData);

                                        string passResetResponse = fai.ChangeUserPass(userData);

                                        if (passResetResponse == "OK")
                                            rowRes = "Ya existe en SAM, se añadió el cliente: " + customer + " - " + passResetResponse;
                                        else
                                            rowRes = "Ya existe en SAM, se añadió el cliente: " + customer + " - Error en cambiar la contraseña:<br>" + passResetResponse;
                                    }
                                    else
                                        rowRes = "Ya existe en SAM, se añadió el cliente: " + customer;
                                }
                                else
                                {
                                    sendIcs = true;
                                    rowRes = fmRes;
                                }

                                #endregion
                            }
                            catch (Exception ex)
                            {
                                sendIcs = true;
                                rowRes = "<b>" + ex.Message + "</b>";
                            }
                        }
                        else
                            rowRes = "Por favor ingrese el Cliente";

                        #endregion
                    }

                    excelRow["Resultado"] = rowRes;
                }

                console.WriteLine("Respondiendo solicitud");

                string htmlTable = val.ConvertDataTableToHTML(excelDt);

                if (sendIcs)
                    mail.SendHTMLMail(htmlTable, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject, new string[] { "smarin@gbm.net", "hlherrera@gbm.net" });
                else
                    mail.SendHTMLMail(htmlTable, new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);

                root.requestDetails = respFinal;

            }
            else
            {
                console.WriteLine("Plantilla incorrecta");
                mail.SendHTMLMail("Utilizar la plantilla oficial de internal customer services", new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);
            }
        }

        private string[] CreateUserInSap(string email, string sapCustomerId, string firstName, string lastName)
        {
            string fmRes, fmPartnerNumber = "", fmMsgResponse = "";

            console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);

            try
            {
                Dictionary<string, string> parameters = new Dictionary<string, string>
                {
                    ["E_MAIL"] = email.ToLower(),
                    ["CUSTOMER"] = sapCustomerId,
                    ["NAME_FIRST"] = firstName,
                    ["NAME_LAST"] = lastName
                };

                IRfcFunction zdmCreateSam = new SapVariants().ExecuteRFC("CRM", "ZDM_CREATE_SAM", parameters);

                fmRes = zdmCreateSam.GetValue("RESULTADO").ToString();
                fmPartnerNumber = zdmCreateSam.GetValue("PARTNER_NUMBER").ToString();
                fmMsgResponse = zdmCreateSam.GetValue("MSGRESPONSE").ToString();

                console.WriteLine(email + ": " + fmRes);

                if (!fmRes.ToLower().Contains("contacto actualizado") && !fmMsgResponse.ToLower().Contains("business partner was created with number") && !fmRes.ToLower().Contains("se agrego el cliente"))
                    fmRes = "<b>Error en SAP, por favor comuníquese con Internal Customer Services</b>";
            }
            catch (Exception ex)
            {
                fmRes = "<b>" + ex.Message + "</b>";
            }

            return new string[] { fmRes, fmPartnerNumber, fmMsgResponse };
        }
    }
}
