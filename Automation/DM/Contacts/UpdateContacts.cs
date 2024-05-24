using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Projects.Contacts;
using DataBotV5.App.Global;
using System.Linq;
using System.Globalization;
using DataBotV5.App.ConsoleApp;

namespace DataBotV5.Automation.DM.Contacts
{    /// <summary><c>ContactsCreation:</c> 
     /// Clase DM Automation encargada de actualización y creación masiva de contactos mediante un e-mail de Smart and Simple en datos maestros.</summary>
    class UpdateContacts
    {
        Log log = new Log();
        CRUD crud = new CRUD();
        Database DB = new Database();
        Rooting root = new Rooting();
        MsExcel MsExcel = new MsExcel();
        ContactSAP cSap = new ContactSAP();
        ValidateData val = new ValidateData();
        ConsoleFormat console = new ConsoleFormat();
        MailInteraction mail = new MailInteraction();
        DeleteContacts delContacts = new DeleteContacts();
        int mandante = 460;
        string respFinal = "";

        public void Main()
        {

            if (mail.GetAttachmentEmail("Solicitudes Modificacion Masiva Contactos", "Procesados", "Procesados Modificacion Masiva Contactos"))
            {
                console.WriteLine("Procesando...");
                string client = "";
                string user = "";
                bool isAutomatic = true;
                if (root.BDUserCreatedBy == "appmanagementsys@gbm.net")
                {
                    //enviado por la pagina de SS
                    user = root.Email_Body.Split(new string[] { "Usuario: " }, 2, StringSplitOptions.None)[1].Split(new char[] { ',' }, 2)[0];
                    client = root.Email_Body.Split(new string[] { "IdCustomer: " }, 2, StringSplitOptions.None)[1].Split(new char[] { ',' }, 2)[0];
                }
                else
                {
                    //enviado por una persona
                    user = root.BDUserCreatedBy; // = "DMEZA@GBM.NET";
                    isAutomatic = false;
                }
                //Start.enviroment = "PRD";
                //root.ExcelFile = "MODIFICAR CONTACTOS.xlsx";

                ProcessUpdate(root.FilesDownloadPath + "\\" + root.ExcelFile, client, user, isAutomatic);

                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }
        /// <summary>
        /// Método convierte el excel en un DataTable, en caso de éxito actualiza en SAP.
        /// </summary>
        /// <param name="ruta">ruta path del archivo del excel</param>
        /// <param name="cliente">el cliente de los contactos a actualizar</param>
        /// <param name="usuario">el usuario que envio la solicitud</param>
        private void ProcessUpdate(string ruta, string cliente, string usuario, bool isAutomatic)
        {
            console.WriteLine("Get Excel...");
            DataTable excel = MsExcel.GetExcel(ruta);
            if (excel == null)
                mail.SendHTMLMail("Error al leer la plantilla de contactos de S&S", new string[] { "internalcustomersrvs@gbm.net" }, "Error al leer la plantilla de contactos de S&S", null);
            else
                UpdateContactsSap(excel, cliente, usuario, isAutomatic);

        }
        /// <summary>
        /// Actualizar contactos en SAP
        /// </summary>
        /// <param name="plantilla"></param>
        /// <param name="cliente"></param>
        /// <param name="usuario"></param>
        private void UpdateContactsSap(DataTable plantilla, string cliente, string usuario, bool isAutomatic)
        {
            string err = "";
            int index = 1;
            bool valContact = false;
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;

            plantilla.Columns.Add("RESULTADO");

            console.WriteLine("Get Data of S&S...");
            DataTable country = crud.Select("SELECT * FROM `Country`", "update_contacts");
            DataTable departament = crud.Select("SELECT * FROM `Departament`", "update_contacts");
            DataTable function = crud.Select("SELECT * FROM `Function`", "update_contacts");

            foreach (DataRow contacto in plantilla.Rows)
            {
                //Diccionario para saber las posibles nombres de columnas del excel ya que viene de S&S pero tambien de una plantilla oficial dada a los usuarios
                Dictionary<string, string[]> columnMappings = new Dictionary<string, string[]>
                {
                    ["cust"] = new string[] { "ID de Cliente", "Customer" },
                    ["idContact"] = new string[] { "ID de Contacto", "Contact" },
                    ["title"] = new string[] { "Tratamiento", "Treatment" },
                    ["country"] = new string[] { "País", "Country" },
                    ["function"] = new string[] { "Funcion", "Function" },
                    ["department"] = new string[] { "Departamento", "Department" },
                    ["language"] = new string[] { "Idioma", "Language" },
                    ["tel1"] = new string[] { "Teléfono de Oficina", "Telephone" },
                    ["mob1"] = new string[] { "Teléfono Celular", "Mobile" },
                    ["tel2"] = new string[] { "Teléfono de Oficina Secundario", "Telephone 2" },
                    ["mob2"] = new string[] { "Teléfono Celular Secundario", "Mobile 2" },
                    ["ext1"] = new string[] { "Extensión", "Ext" },
                    ["ext2"] = new string[] { "Extensión Secundaria", "Ext2" },
                    ["name"] = new string[] { "Nombre", "Name" },
                    ["lname"] = new string[] { "Apellido", "Last Name" },
                    ["addr"] = new string[] { "Dirección", "Address" },
                    ["email"] = new string[] { "Correo electrónico", "Email" },
                    ["delete"] = new string[] { "Eliminar Contacto", "Delete" },
                    ["sustituto"] = new string[] { "Contacto Sustituto", "Contacto Sustituto" },
                    ["confirm"] = new string[] { "Usuario Confirmador", "Confirm Responsable" },
                };

                string contacto_id = "";
                try
                {
                    if (!isAutomatic) cliente = GetColumnValue(contacto, columnMappings["cust"]).PadLeft(10, '0');

                    #region Arreglar campos

                    foreach (DataColumn column in plantilla.Columns)
                    {
                        if (contacto[column].ToString() == "N/A" || contacto[column].ToString() == "NA")
                        {
                            contacto[column] = "";
                        }
                    }

                    ContactInfo contactInfo = new ContactInfo();

                    contacto_id = GetColumnValue(contacto, columnMappings["idContact"]); //  contacto["ID de Contacto"].ToString(); //****

                    if (!string.IsNullOrWhiteSpace(contacto_id))
                    {
                        contacto_id = contacto_id.PadLeft(10, '0');
                    }
               
                    //Checkea si hay que eliminarlo
                    string delContact = GetColumnValue(contacto, columnMappings["delete"]);
                    string sustituteContact = GetColumnValue(contacto, columnMappings["sustituto"]);

                    if (delContact.ToLower() == "x" && sustituteContact.Trim() != "")
                    {
                        //ver como procesar el tema de eliminado ya que se necesita un sustituto si es que el contacto tiene documentos asociados
                        DataTable dt = new DataTable();
                        dt.Columns.Add("idContactLock");
                        dt.Columns.Add("idContactSubstitute");
                        dt.Columns.Add("idCustomer");
                        dt.Columns.Add("userName");

                        DataRow dr = dt.Rows.Add();
                        dr["idContactLock"] = contacto_id;
                        dr["idContactSubstitute"] = sustituteContact;
                        dr["idCustomer"] = cliente;
                        dr["userName"] = GetColumnValue(contacto, columnMappings["confirm"]); ;

                        dt.AcceptChanges();
                        string contactDelResponse = delContacts.deleteContactsVoid(dt, false);
                        contacto["RESULTADO"] = contactDelResponse;
                        continue;
                    }

                    string tratamiento = GetColumnValue(contacto, columnMappings["title"]);  //contacto["Tratamiento"].ToString(); //****
                    string pais = GetColumnValue(contacto, columnMappings["country"]); // contacto["País"].ToString();//****
                    string funcion = GetColumnValue(contacto, columnMappings["function"]);  //contacto["Funcion"].ToString();//****
                    string departamento = GetColumnValue(contacto, columnMappings["department"]);  //contacto["Departamento"].ToString();//****
                    string idioma = GetColumnValue(contacto, columnMappings["language"]);  //contacto["Idioma"].ToString();//****

                    try { pais = country.Select("Name_Country = '" + textInfo.ToTitleCase(pais) + "'")[0]["Code_Country"].ToString(); }
                    catch
                    {
                        try
                        {
                            pais = country.Select("Code_Country = '" + pais + "'")[0]["Code_Country"].ToString();
                        }
                        catch
                        {

                            if (pais != "")
                            {

                                contacto["RESULTADO"] = "Ingrese un país correcto";
                                continue;
                            }
                        }

                    }//****
                    try { departamento = departament.Select("Name_Departament = '" + textInfo.ToTitleCase(departamento) + "'")[0]["Code_Departament"].ToString(); }
                    catch
                    {
                        contacto["RESULTADO"] = "Ingrese un Departamento correcto";
                        continue;
                    }//****
                    try { funcion = function.Select("Name_Function = '" + textInfo.ToTitleCase(funcion) + "'")[0]["Code_Function"].ToString(); }
                    catch
                    {
                        if (funcion != "")
                        {
                            contacto["RESULTADO"] = "Ingrese una Funcion correcta";
                            continue;
                        }
                    }//****

                    if (tratamiento.ToLower() == "señora")
                        tratamiento = "0001";
                    else if (tratamiento.ToLower() == "señor")
                        tratamiento = "0002";
                    else if (tratamiento.Trim() == "")
                    {
                        tratamiento = "";
                    }
                    else
                    {
                        contacto["RESULTADO"] = "Ingrese un tratamiento correcto";
                        continue;
                    }



                    if (idioma.ToLower() == "inglés" || idioma.ToLower() == "ingles" || idioma.ToLower() == "e" || idioma.ToLower() == "en")
                        idioma = "EN";
                    else if (idioma.ToLower() == "español" || idioma.ToLower() == "espanol" || idioma.ToLower() == "s" || idioma.ToLower() == "es")
                        idioma = "ES";
                    else if (idioma == "")
                    {

                    }
                    else
                    {
                        contacto["RESULTADO"] = "Ingrese un lenguaje correcto";
                        continue;
                    }


                    List<phones> telefonos = new List<phones>();
                    string tel1 = GetColumnValue(contacto, columnMappings["tel1"]); //contacto["Teléfono de Oficina"].ToString();//****
                    string mob1 = GetColumnValue(contacto, columnMappings["mob1"]); //contacto["Teléfono Celular"].ToString();//****
                    string ext1 = GetColumnValue(contacto, columnMappings["ext1"]); //contacto["Extensión"].ToString();//****

                    string tel2 = GetColumnValue(contacto, columnMappings["tel2"]); //contacto["Teléfono de Oficina Secundario"].ToString();//****
                    string mob2 = GetColumnValue(contacto, columnMappings["mob2"]); //contacto["Teléfono Celular Secundario"].ToString();//****
                    string ext2 = GetColumnValue(contacto, columnMappings["ext2"]); //contacto["Extensión Secundaria"].ToString();//****

                    if (!string.IsNullOrWhiteSpace(tel1) || !string.IsNullOrWhiteSpace(mob1) || !string.IsNullOrWhiteSpace(ext1))
                    {
                        phones phone = new phones();
                        phone.TELEPHONE = tel1;
                        phone.MOBILE = mob1;
                        phone.EXT = ext1;
                        telefonos.Add(phone);
                    }

                    if (!string.IsNullOrWhiteSpace(tel2) || !string.IsNullOrWhiteSpace(mob2) || !string.IsNullOrWhiteSpace(ext2))
                    {
                        phones phone2 = new phones();
                        phone2.TELEPHONE = tel2;
                        phone2.MOBILE = mob2;
                        phone2.EXT = ext2;
                        telefonos.Add(phone2);
                    }


                    #endregion

                    contactInfo.cliente = cliente;
                    contactInfo.contacto = contacto_id;
                    contactInfo.tratamiento = tratamiento;
                    contactInfo.nombre = GetColumnValue(contacto, columnMappings["name"]); //contacto["Nombre"].ToString();//****
                    contactInfo.apellido = GetColumnValue(contacto, columnMappings["lname"]); //contacto["Apellido"].ToString();//****
                    contactInfo.pais = pais;
                    contactInfo.direccion = GetColumnValue(contacto, columnMappings["addr"]); //contacto["Dirección"].ToString().ToUpper();//****
                    contactInfo.email = GetColumnValue(contacto, columnMappings["email"]); //contacto["Correo electrónico"].ToString();//****
                    contactInfo.telefonos = telefonos;
                    contactInfo.puesto = funcion;
                    contactInfo.departamento = departamento;
                    contactInfo.idioma = idioma;

                    #region FM actualizar contacto
                    string resp = "";
                    string tipo = "";
                    if (string.IsNullOrWhiteSpace(contacto_id))
                    {
                        resp = cSap.CreateContactSAP(contactInfo);
                        tipo = "C";
                    }
                    else
                    {
                        resp = cSap.UpdateContactSAP(contactInfo);
                        tipo = "M";
                    }

                    #endregion

                    #region Procesar Resultados

                    if (resp.Contains("Error"))
                    {
                        valContact = true;
                        err = err + resp + "<br>";
                    }
                    else
                    {
                        if (tipo == "C")
                        {
                            try
                            {
                                contacto_id = resp.Split(new string[] { "ID: " }, 2, StringSplitOptions.None)[1].ToString();

                            }
                            catch (Exception)
                            {
                                contacto_id = resp.Split(new string[] { ": " }, 2, StringSplitOptions.None)[1].ToString();
                                contacto_id = contacto_id.Replace(": contacto ya existe<br>", "");

                            }
                        }
                        contacto_id = contacto_id.Replace("<br>", "");
                        contacto_id = contacto_id.PadLeft(10, '0');

                        //if (isAutomatic)
                        //{
                            resp = "{" + $"\"ID_CUSTOMER\": \"{cliente}\",\"ID_CONTACT_CRM\": \"{contacto_id}\"" + "}";
                            string sql1 = "INSERT INTO `HistoryContacts` (`Id_Customer`, `Id_Contact`, `Create_By` ,`Update_By`, `Change_Values`, `Type_Data`) " +
                                "VALUES ('" + cliente + "','" + contacto_id + "', '" + usuario + "', '" + usuario + "', '" + resp.Replace("<br>", "") + $"', '{tipo}')";

                            crud.Insert(sql1, "update_contacts");
                        //}


                        //Confirmar el contacto como revisado
                        string sqlSelect = $"SELECT * FROM `ConfirmContacts` WHERE `idCustomer` = '{cliente}' AND `idContact` = '{contacto_id}'";
                        DataTable respConfirm = crud.Select(sqlSelect, "update_contacts");
                        if (respConfirm.Rows.Count <= 0) //no esta confirmado
                        {

                            string createdBy = usuario;
                            if (!isAutomatic)
                            {
                                //si es enviado por una persona normal se le cae encima al usuario 
                                createdBy = GetColumnValue(contacto, columnMappings["confirm"]);
                                if (createdBy == "")
                                {
                                    ////sacar el employee responsable del cliente:
                                    string sqlSelectAM = $"SELECT accountManagerUser FROM `clients` WHERE `idClient` = '{cliente.TrimStart(new char[] { '0' })}'";
                                    DataTable resp2 = crud.Select(sqlSelectAM, "databot_db");
                                    if (resp2.Rows.Count >= 0)
                                    {
                                        createdBy = resp2.Rows[0]["accountManagerUser"].ToString();
                                    }
                                    if (createdBy == "")
                                    {
                                        //toma el usuario que mando el correo
                                        createdBy = root.BDUserCreatedBy;
                                    }
                                }
                            }



                            string sqlConfirm = $@"INSERT INTO ConfirmContacts (idClient, idCustomer, idContact, createdBy)
                                            SELECT 
                                                c.id, 
                                                '{cliente}',
                                                '{contacto_id}',
                                                '{createdBy.ToUpper()}'
                                            FROM 
                                                (SELECT '{cliente.TrimStart(new char[] { '0' })}' AS idClient) AS dummy
                                            LEFT JOIN 
                                                databot_db.clients AS c ON c.idClient = '{cliente.TrimStart(new char[] { '0' })}';";

                            crud.Insert(sqlConfirm, "update_contacts");
                        }

                        //la respuesta final

                        resp = $"Contacto {((tipo == "C") ? "creado" : "modificado")} con éxito: {contacto_id}";

                        //log de ICS
                        log.LogDeCambios(tipo, root.BDProcess, usuario + "(S&S)", "Actualizar contacto", resp, "");
                        respFinal = respFinal + "\\n" + $"Actualizar contacto {contacto_id}: " + resp;

                    }

                    //si llega hasta aqui todo ok

                    contacto["RESULTADO"] = resp;
                    if (tipo == "C")
                    {
                        contacto[GetColumnName(contacto, columnMappings["idContact"])] = contacto_id;
                    }

                    #endregion
                }
                catch (Exception ex)
                {
                    string msjEx = "Error al actualizar contacto: " + contacto_id + " en la linea: " + index + "<br><br>" + ex.ToString();
                    console.WriteLine(msjEx);
                    contacto["RESULTADO"] = msjEx;
                    err = err + msjEx;
                    valContact = true;
                }
                index++;
            }

            if (valContact)
            {
                mail.SendHTMLMail("Error en actualizacion de contactos masivos de S&S<br>" + err + "<br>Solicitud:<br><br>" + val.ConvertDataTableToHTML(plantilla),
                      new string[] { "internalcustomersrvs@gbm.net", "appmanagement@gbm.net", "dmeza@gbm.net" }, $"Error en actualizacion de contactos masivos de S&S {usuario} - {cliente}", new string[] { "joarojas@gbm.net" });

            }
            console.WriteLine("Save Excel...");
            string ruta = root.FilesDownloadPath + $"\\Plantilla Resultados {cliente}.xlsx";
            MsExcel.CreateExcel(plantilla, "Resultados", ruta);
            string msj = "Estimado(a) se le notifica que los contactos del cliente: " + cliente + " se han actualizado de acuerdo al excel adjunto <br><br> Por favor verifique refrescando en Smart & Simple";
            string html = Properties.Resources.emailtemplate1;
            html = html.Replace("{subject}", "Creación/Actualización Masiva de Contactos");
            html = html.Replace("{cuerpo}", msj);
            html = html.Replace("{contenido}", "");
            console.WriteLine("Send Email...");
            string sender = (isAutomatic) ? usuario + "@GBM.NET" : root.BDUserCreatedBy;

            mail.SendHTMLMail(html, new string[] { sender }, $"Notificacion solicitud actualización de contactos del cliente {cliente}", root.CopyCC, new string[] { ruta });

            root.requestDetails = respFinal;

        }
        public string GetColumnValue(DataRow row, params string[] columnNames)
        {
            var column = columnNames.FirstOrDefault(name => row.Table.Columns.Contains(name));
            return column != null ? row[column].ToString().ToUpper().Trim() : "";
        }
        public string GetColumnName(DataRow row, params string[] columnNames)

        {
            var column = columnNames.FirstOrDefault(name => row.Table.Columns.Contains(name));
            return column;
        }
    }

}
