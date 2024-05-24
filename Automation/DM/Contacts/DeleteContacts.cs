using SAP.Middleware.Connector;
using System;
using System.Data;
using System.Linq;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Logical.Webex;
using DataBotV5.App.Global;
using DataBotV5.Data.Stats;

namespace DataBotV5.Automation.DM.Contacts
{    /// <summary><c>ContactsCreation:</c> 
     /// Clase DM Automation encargada de eliminación de contactos.</summary>
    class DeleteContacts
    {
        ConsoleFormat console = new ConsoleFormat();
        Rooting root = new Rooting();
        MailInteraction mail = new MailInteraction();

        WebexTeams wt = new WebexTeams();
        SapVariants sap = new SapVariants();
        CRUD crud = new CRUD();
        Log log = new Log();

        string respFinal = "";



        string mandante = "CRM";

        public void Main()
        {
            DataTable del_contacts = null;
            string filas, resSustituto, resBloqueo, mensaje, resError;
            filas = resSustituto = resBloqueo = resError = mensaje = "";
            bool valError = false;

            #region Leer solicitudes
            try
            {
                string sql3 = "SELECT * FROM `LogicLock` WHERE status = 1";
                del_contacts = crud.Select(sql3, "update_contacts");
            }
            catch (Exception ex)
            {
                try
                {
                    crud.Update($"UPDATE `orchestrator` SET `active`= 0 WHERE `class` = '{root.BDMethod}'", "databot_db");

                }
                catch (Exception ex2)
                {

                }
                mail.SendHTMLMail("No se pudo conectar con la BD de S&S<br><br>" + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, "Error en eliminar contactos de S&S<br>", new string[] { "dmeza@gbm.net", "joarojas@gbm.net" });
            }
            #endregion

            if (del_contacts.Rows.Count > 0)
            {
                deleteContactsVoid(del_contacts, true);

                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }

        }
        
        public string deleteContactsVoid(DataTable del_contacts, bool fromSS)
        {
            string response = "";
            string filas, resSustituto, resBloqueo, mensaje, resError;
            filas = resSustituto = resBloqueo = resError = mensaje = "";
            bool valError = false;
            console.WriteLine("Procesando...");
            del_contacts.Columns.Add("RESULTADO");

            DataTable distinct_users = del_contacts.DefaultView.ToTable(true, "userName");

            foreach (DataRow user in distinct_users.Rows)
            {
                filas = resSustituto = resBloqueo = resError = mensaje = "";
                console.WriteLine($"Enviando solicitudes del usuario: {user["userName"]}");
                foreach (DataRow contacto in del_contacts.Select("userName = '" + user["userName"] + "'"))   //hacer usuario por usuario
                {
                    #region FM leer documentos asociados a contacto

                    string old_contact = contacto["idContactLock"].ToString();
                    try
                    {
                        string new_contact = contacto["idContactSubstitute"].ToString();
                        string cliente = contacto["idCustomer"].ToString();

                        if (fromSS)
                        {
                            filas += contacto["Id"].ToString() + ",";
                        }


                        RfcDestination dest_crm = sap.GetDestRFC(mandante);
                        IRfcFunction get_documents = dest_crm.Repository.CreateFunction("ZGET_ASOC_DOCUMENTS");
                        get_documents.SetValue("PARTNER", old_contact);
                        get_documents.SetValue("PARTNER_FCT", "00000015");
                        get_documents.SetValue("LANGUAGE", "EN");

                        get_documents.Invoke(dest_crm);


                        #endregion

                        #region Cambiar de contacto los documentos encontrados
                        DataTable info = sap.GetDataTableFromRFCTable(get_documents.GetTable("INFO"));

                        string documentos = "";
                        resBloqueo = "";
                        bool delete = true;
                        string[] status = { "Open", "In Process", "Being Processed by Customer", "Waiting for customer",
                                "Waiting for partner", "In progress", "In Review HQ", "On Hold" };
                        if (info.Rows.Count > 0)
                        {

                            foreach (DataRow documento in info.Rows)
                            {

                                if (status.Contains(documento["STATUS"].ToString()))
                                {

                                    IRfcFunction asoc_documents = dest_crm.Repository.CreateFunction("ZPUT_ASOC_DOCUMENTS_CRM");

                                    asoc_documents.SetValue("NEW_PARTNER", new_contact); //devuelve OK aunque no exista
                                    asoc_documents.SetValue("PARTNER_FCT", "00000015");
                                    asoc_documents.SetValue("DOCUMENT_ID", documento["ID"].ToString());
                                    asoc_documents.SetValue("DOCUMENT_TYPE", documento["TYPE"].ToString());
                                    asoc_documents.Invoke(dest_crm);


                                    resSustituto = asoc_documents.GetValue("RESPONSE").ToString();

                                    if (resSustituto != "OK")
                                    {
                                        delete = false;
                                        documentos += $"- {documento["ID"].ToString()} - {resSustituto} \r\n";
                                    }
                                }
                            }
                        }
                        #endregion

                        #region Eliminar relacion del contacto viejo
                        if (delete)
                        {
                            IRfcFunction del_rel = dest_crm.Repository.CreateFunction("ZICS_BP_DEL_CONTACT");
                            del_rel.SetValue("CLIENTE", cliente.TrimStart(new char[] { '0' }));
                            del_rel.SetValue("CONTACTO", old_contact.TrimStart(new char[] { '0' }));
                            del_rel.Invoke(dest_crm);



                            resBloqueo = del_rel.GetValue("RESPUESTA").ToString();
                        }

                        #endregion

                        #region Procesar resultados
                        if (resBloqueo == "OK")
                        {
                            response = resBloqueo + " ID: " + contacto["idContactLock"].ToString();
                            contacto["RESULTADO"] = response;
                            //modificar el status de la solicitud
                            if (fromSS)
                            {

                                string sqlUpdate = $"UPDATE `LogicLock` SET `status` = '0', `updatedAt` = '{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}' WHERE `Id` = " + contacto["Id"].ToString();
                                crud.Update(sqlUpdate, "update_contacts");

                                mensaje += $"Se reemplazo el contacto: {old_contact} por el contacto sustituto: {new_contact}, del cliente: {cliente}\n\n";

                                //Desconfirmar el contacto como revisado
                                //string sqlSelect = $"SELECT * FROM `ConfirmContacts` WHERE `idCustomer` = '{cliente}' AND `idContact` = '{old_contact}'";
                                //DataTable resp = crud.Select(sqlSelect, "update_contacts");

                                //if (resp.Rows.Count > 0) //esta confirmado
                                //{
                                //    try
                                //    {
                                //        string sqlDelete = $"DELETE FROM `ConfirmContacts` WHERE `idCustomer` = '{cliente}' AND `idContact` = '{old_contact}'";
                                //        crud.Delete(sqlDelete, "update_contacts");

                                //        string respo = $"Eliminar el contacto: {old_contact}, del cliente: {cliente}";
                                //        log.LogDeCambios("Eliminar", root.BDProcess, new string[] { root.BDUserCreatedBy }, "Eliminar contacto DM", respo, root.Subject);
                                //        respFinal = respFinal + "\\n" + respo;


                                //    }
                                //    catch (Exception EX)
                                //    {
                                //        mail.SendHTMLMail($"Error en eliminacion de contactos de S&S a la hora de desconfirmar el contacto {old_contact} del cliente {cliente}<br>" + EX,
                                //         new string[] { "internalcustomersrvs@gbm.net", new string[] {"appmanagement@gbm.net"}, new string[] { "dmeza@gbm.net" }, "joarojas@gbm.net" }, $"Error en eliminacion de contactos de S&S {user["userName"]}", 2);
                                //    }
                                //}

                            }
                        }
                        else if (!delete)
                        {

                            if (fromSS)
                            {

                                //Confirmar el contacto como revisado
                                string sqlSelect = $"SELECT * FROM `ConfirmContacts` WHERE `idCustomer` = '{cliente}' AND `idContact` = '{old_contact}'";
                                DataTable resp = crud.Select(sqlSelect, "update_contacts");

                                if (resp.Rows.Count <= 0) //no esta confirmado
                                {
                                    //sacar el usuario del cliente:
                                    string sqlSelectAM = $"SELECT accountManagerUser FROM `clients` WHERE `idClient` = '{cliente.TrimStart(new char[] { '0' })}'";
                                    DataTable resp2 = crud.Select(sqlSelectAM, "databot_db");
                                    if (resp2.Rows.Count > 0)
                                    {
                                        user["userName"] = resp2.Rows[0]["accountManagerUser"].ToString();
                                    }

                                    string sqlConfirm = "INSERT INTO `ConfirmContacts` (`idCustomer`, `idContact`, `CreateBy`) " +
                                        $"VALUES ('{cliente}','{old_contact}', '{user["userName"].ToString()}')";

                                    crud.Insert(sqlConfirm, "update_contacts");
                                }
                            }

                            mensaje = mensaje + "No se pudo eliminar el contacto: " + old_contact + " debido a que no se pudo sustituir el contacto en los siguientes documentos de venta \r\n" + documentos + "\n\n";
                            response = mensaje;
                        }
                        else//error
                        {
                            mensaje = mensaje + "No se pudó eliminar el contacto: " + old_contact + " debido a un error en SAP " + resBloqueo + " \n";
                            string sqlU = $"UPDATE `LogicLock` SET `status` = '2', `updatedAt` = '{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}' WHERE `Id` = " + contacto["Id"].ToString();
                            crud.Update(sqlU, "update_contacts");
                            response = resBloqueo + " ID: " + contacto["Id"].ToString();
                            valError = true;
                            contacto["RESULTADO"] = response;
                        }
                        #endregion
                    }
                    catch (Exception ex)
                    {
                        mensaje = mensaje + "No se pudó eliminar el contacto: " + old_contact + " debido a un error en SAP " + ex.ToString() + " \n";
                        string sqlU = $"UPDATE `LogicLock` SET `status` = '2', `updatedAt` = '{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}' WHERE `Id` = " + contacto["Id"].ToString();
                        crud.Update(sqlU, "update_contacts");
                        response = contacto["idContactLock"].ToString() + "<br>" + ex.ToString();
                        contacto["RESULTADO"] = response;
                        valError = true;
                    }
                } //foreach

                if (valError)
                {
                    string msj = mensaje.Replace(" \n\n", "<br>");
                    msj = msj.Replace(" \r\n", "<br>");
                    msj = msj.Replace(" \n", "<br>");
                    mail.SendHTMLMail("Error en eliminacion de contactos de S&S<br>" + msj,
                        new string[] { "internalcustomersrvs@gbm.net", "appmanagement@gbm.net", "dmeza@gbm.net", "joarojas@gbm.net" }, "Error en eliminacion de contactos de S&S");

                }
                console.WriteLine("Enviando Notificaciones...");

                #region Enviar notificacion al solicitante
                if (fromSS)
                {

                    wt.SendNotification(user["userName"] + "@GBM.NET", "Solicitud deshabilitar contacto",
                    "**Notificacion Bloqueo de Contacto:** Estimado(a) se le notifica que sus solicitudes han dado el siguiente resultado: <br><br>" + mensaje);
                }
                #endregion


                root.requestDetails = respFinal;
                root.BDUserCreatedBy = user["userName"].ToString();


            }
            return response;

        }



    }

}
