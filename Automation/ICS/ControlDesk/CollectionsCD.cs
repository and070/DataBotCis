using DataBotV5.Logical.Projects.ControlDesk;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using System.Data;
using System.Linq;
using System;

namespace DataBotV5.Automation.ICS.ControlDesk
{
    internal class CollectionsCD
    {
        readonly ControlDeskInteraction cdi = new ControlDeskInteraction();
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly ValidateData val = new ValidateData();
        readonly Credentials cred = new Credentials();
        readonly MsExcel excel = new MsExcel();
        readonly Rooting root = new Rooting();
        readonly Log log = new Log();

        const string mandCd = "QAS";
        public void Main()
        {
            cred.SelectCdMand(mandCd);

            //Creación
            mail.GetAttachmentEmail("Solicitudes CD Collections", "Procesados", "Procesados CD Collections");
            if (root.ExcelFile != "")
            {
                string filePath = root.FilesDownloadPath + "\\" + root.ExcelFile;
                DataTable excelDt = excel.GetExcelBook(filePath).Tables["Formulario Colecciones"];
                CreateCollections(excelDt, filePath);
                using (Stats stats = new Stats()) { stats.CreateStat(); }
            }

            //Modificación
            mail.GetAttachmentEmail("Solicitudes CD Collections Mod", "Procesados", "Procesados CD Collections");
            if (root.ExcelFile != "")
            {
                string filePath = root.FilesDownloadPath + "\\" + root.ExcelFile;
                DataTable excelDt = excel.GetExcelBook(filePath).Tables["Formulario Colecciones"];
                ChangeCollections(excelDt, filePath);
                using (Stats stats = new Stats()) { stats.CreateStat(); }
            }
        }

        private void ChangeCollections(DataTable excelDt, string filePath)
        {
            DataTable resDt = new DataTable();
            resDt.Columns.Add("Collection ID");
            resDt.Columns.Add("User");
            resDt.Columns.Add("Type");
            resDt.Columns.Add("Action");
            resDt.Columns.Add("Result");
            resDt.Columns.Add("Supervisor");

            if (excelDt != null)
            {
                for (int i = 15; i < excelDt.Rows.Count; i++) //15 es la fila donde empiezan los IDs
                {
                    DataRow resRow = resDt.NewRow();
                    DataRow row = excelDt.Rows[i];
                    string colNum = row[0].ToString();
                    resRow[0] = colNum;
                    resRow[5] = row[2].ToString();
                    CdCollectionData collectionInfo = cdi.GetCollectionData(colNum);

                    List<string> typeList = new List<string>();
                    if (collectionInfo.CollectionNum != null)
                    {
                        string action = row[9].ToString();
                        string contact = row[3].ToString();

                        if (action.ToUpper() == "ADD")
                        {
                            foreach (CdCollectionPartyData party in collectionInfo.CollectionParties)
                            {
                                if (row[3].ToString().ToUpper() == party.PersonId) //buscar el user en la colección
                                    typeList.Add(party.Type);

                                if (row[3].ToString().ToUpper() == party.PersonGroup) //buscar el person group en la colección
                                    typeList.Add(party.Type);
                            }

                            if (typeList.Count != 0) //el usuario ya existe en la colección
                            {
                                if (typeList.Contains(row[7].ToString()))
                                {
                                    //ya existe el usuario con el type
                                    resRow[4] = "El usuario o grupo ya existe en la colección";
                                }
                                else //el usuario ya existe pero sin el type
                                {
                                    CdCollectionPartyData partyData = new CdCollectionPartyData
                                    {
                                        Description = row[8].ToString(),
                                        Type = row[7].ToString()
                                    };

                                    if (contact.Contains("@")) //es correo
                                    {
                                        string email = contact;
                                        contact = string.Concat(contact.Take(30));

                                        string personId = cdi.CheckPersonExistence(email);

                                        #region Crear contacto
                                        if (personId == "NE")
                                        {
                                            string phone = row[6].ToString();

                                            if (!phone.Contains('+'))
                                                phone = "+" + phone;

                                            phone = val.ParsePhoneNumberToE164(phone);

                                            if (!phone.Contains("ERROR"))
                                            {
                                                CdContactData personData = new CdContactData
                                                {
                                                    PersonId = contact,
                                                    FirstName = row[4].ToString(),
                                                    LastName = row[5].ToString(),
                                                    Email = email,
                                                    Telephone = phone
                                                };

                                                string contactMsg = cdi.CreateContact(personData);

                                                if (contactMsg != "OK")
                                                {
                                                    resRow[4] += "Error contacto: " + contactMsg;
                                                    console.WriteLine(resRow[0] + ": Error contacto: " + contactMsg);
                                                }
                                            }
                                            else
                                            {
                                                resRow[4] += phone;
                                                console.WriteLine(resRow[0] + ": Error contacto: " + phone);
                                            }
                                        }
                                        #endregion

                                        partyData.PersonId = contact.ToUpper();
                                    }
                                    else //es Grupo
                                    {
                                        string personGroup = cdi.CheckPersonGroupExistence(contact.ToUpper());

                                        if (personGroup != "NE")
                                            partyData.PersonGroup = contact.ToUpper();
                                        else
                                            resRow[4] = "El grupo no existe";
                                    }

                                    //si no hay errores aplique cambios en CD
                                    if (resRow[4].ToString().Length == 0)
                                        resRow[4] = cdi.AddCollectionParty(partyData, colNum);
                                }
                            }
                            else //el usuario no existe el usuario en la colección
                            {
                                CdCollectionPartyData partyData = new CdCollectionPartyData
                                {
                                    Description = row[8].ToString(),
                                    Type = row[7].ToString()
                                };

                                if (contact.Contains("@"))
                                {
                                    if (contact.ToLower().Contains("@gbm.net"))
                                    {
                                        string email = contact;
                                        contact = string.Concat(contact.Take(30));

                                        string personId = cdi.CheckPersonExistence(email);

                                        if (personId == "NE")
                                            resRow[4] = "El contacto GBM no existe";
                                        else
                                            partyData.PersonId = contact.ToUpper();

                                    }
                                    else
                                    {
                                        string email = contact;
                                        contact = string.Concat(contact.Take(30));

                                        string personId = cdi.CheckPersonExistence(email);

                                        #region Crear contacto
                                        if (personId == "NE")
                                        {
                                            string phone = row[6].ToString();

                                            if (!phone.Contains('+'))
                                                phone = "+" + phone;

                                            phone = val.ParsePhoneNumberToE164(phone);

                                            if (!phone.Contains("ERROR"))
                                            {
                                                CdContactData personData = new CdContactData
                                                {
                                                    PersonId = contact,
                                                    FirstName = row[4].ToString(),
                                                    LastName = row[5].ToString(),
                                                    Email = email,
                                                    Telephone = phone
                                                };

                                                string contactMsg = cdi.CreateContact(personData);

                                                if (contactMsg != "OK")
                                                {
                                                    resRow[4] += "Error contacto: " + contactMsg;
                                                    console.WriteLine(resRow[0] + ": Error contacto: " + contactMsg);
                                                }
                                            }
                                            else
                                            {
                                                resRow[4] += phone;
                                                console.WriteLine(resRow[0] + ": Error contacto: " + phone);
                                            }
                                        }
                                        #endregion

                                        partyData.PersonId = contact.ToUpper();
                                    }
                                }
                                else
                                {
                                    string personGroup = cdi.CheckPersonGroupExistence(contact.ToUpper());

                                    if (personGroup != "NE")
                                        partyData.PersonGroup = contact.ToUpper();
                                    else
                                        resRow[4] = "El grupo no existe";
                                }

                                //si no hay errores aplique cambios en CD
                                if (resRow[4].ToString().Length == 0)
                                    resRow[4] = cdi.AddCollectionParty(partyData, colNum);
                            }
                        }
                        else if (action.ToUpper() == "DELETE")
                        {
                            List<string> idsList = new List<string>();

                            foreach (CdCollectionPartyData party in collectionInfo.CollectionParties)
                            {
                                if (row[3].ToString().ToUpper() == party.PersonId && row[7].ToString().ToUpper() == party.Type.ToUpper()) //buscar el user en la coleccion
                                    idsList.Add(party.Id);

                                if (row[3].ToString().ToUpper() == party.PersonGroup && row[7].ToString().ToUpper() == party.Type.ToUpper()) //buscar el person group en la coleccion
                                    idsList.Add(party.Id);
                            }

                            if (idsList.Count != 0)
                            {
                                //TOMAR SU id
                                foreach (string id in idsList)
                                {
                                    //ELIMINAR ESE ID con el metodo DeleteCollectionParty
                                    string actionRes = cdi.DeleteCollectionParty(id); ;
                                    resRow[4] = actionRes;
                                }
                            }
                            else
                                resRow[4] = "La relación contacto-tipo no existe en la colección";

                        }

                        resRow[0] = colNum;
                        resRow[1] = contact.ToUpper();
                        resRow[2] = row[7].ToString().ToUpper();
                        resRow[3] = action;
                        resDt.Rows.Add(resRow);

                        //LOG
                        string logMsg = $"{resRow[0]} | {resRow[1]} | {resRow[2]} | {resRow[3]} | {resRow[4]}";
                        log.LogDeCambios("", "", root.BDUserCreatedBy, "Modificación Colección en CD", logMsg, "");
                        root.requestDetails += "Modificación Colección en CD: " + logMsg;
                        console.WriteLine("Modificación Colección en CD: " + logMsg);
                    }
                }

                //EMAILS
                console.WriteLine("Enviando correo de OK");
                SendResponseMail("Resultado de modificación de colecciones en Control desk<br><br>", resDt, new string[] { filePath });
            }
            else
            {
                console.WriteLine("Enviando correo de plantilla errónea");
                mail.SendHTMLMail("La plantilla no contiene ninguna hoja con el nombre \"Formulario Colecciones\"", new string[] { root.BDUserCreatedBy }, root.BDClass, attachments: new string[] { filePath });
            }
        }
        private void CreateCollections(DataTable excelDt, string filePath)
        {
            bool sendIcs = false;
            DataTable resDt = new DataTable();
            resDt.Columns.Add("Collection ID");
            resDt.Columns.Add("Supervisor");
            resDt.Columns.Add("Descripción");
            resDt.Columns.Add("Respuesta");

            if (excelDt != null)
            {
                string col0 = excelDt.Columns[0].ColumnName;
                DataTable distinctCollections = new DataView(excelDt).ToTable(true, new string[] { col0 });

                for (int i = 7; i < distinctCollections.Rows.Count; i++) //7 es la fila donde empiezan los IDs
                {
                    CdCollectionData col = new CdCollectionData();
                    DataRow resRow = resDt.NewRow();
                    DataRow row = distinctCollections.Rows[i];
                    string colNum = row[0].ToString();
                    resRow[0] = colNum;

                    if (!cdi.CheckCollectionExistence(colNum))
                    {
                        try
                        {
                            List<CdCollectionPartyData> colParties = new List<CdCollectionPartyData>();

                            DataRow[] singleCollection = excelDt.Select(col0 + " = '" + colNum + "'");
                            string colDesc = singleCollection[0][1].ToString();
                            colDesc = string.Concat(colDesc.Take(100));

                            foreach (DataRow item in singleCollection)
                            {
                                CdCollectionPartyData colParty = new CdCollectionPartyData();

                                string contact = item[3].ToString();
                                string email = item[3].ToString();

                                if (contact.Contains("@"))
                                {
                                    if (contact.ToLower().Contains("@gbm.net"))
                                    {
                                        contact = string.Concat(contact.Take(30));

                                        string personId = cdi.CheckPersonExistence(email);

                                        if (personId == "NE")
                                            resRow[3] += "El contacto GBM no existe\n";
                                        else
                                            colParty.PersonId = contact.ToUpper();

                                    }
                                    else
                                    {
                                        contact = string.Concat(contact.Take(30));

                                        string personId = cdi.CheckPersonExistence(email);

                                        #region Crear contacto
                                        if (personId == "NE")
                                        {
                                            string phone = item[6].ToString();

                                            if (!phone.Contains('+'))
                                                phone = "+" + phone;

                                            if (!phone.Contains("ERROR"))
                                            {
                                                CdContactData personData = new CdContactData
                                                {
                                                    PersonId = contact,
                                                    FirstName = item[4].ToString(),
                                                    LastName = item[5].ToString(),
                                                    Email = email,
                                                    Telephone = phone
                                                };

                                                string contactMsg = cdi.CreateContact(personData);

                                                if (contactMsg != "OK")
                                                {
                                                    resRow[3] += "Error contacto: " + contactMsg + "\n";
                                                    console.WriteLine(resRow[0] + ": Error contacto: " + contactMsg);
                                                }
                                            }
                                            else
                                            {
                                                resRow[3] += phone + "\n";
                                                console.WriteLine(resRow[0] + ": " + phone);
                                            }
                                        }
                                        #endregion

                                        colParty.PersonId = contact;

                                    }
                                }
                                else
                                    colParty.PersonGroup = contact;

                                //llenar los colparties
                                string partyDesc = item[8].ToString();
                                partyDesc = string.Concat(partyDesc.Take(256));

                                colParty.Type = item[7].ToString();
                                colParty.Description = partyDesc;

                                colParties.Add(colParty);
                            }

                            col.CollectionNum = colNum;
                            col.Description = colDesc;
                            col.Supervisor = singleCollection[0][2].ToString();
                            col.CollectionParties = colParties;

                            string collMsg = cdi.CreateCollection(col);

                            if (collMsg != "OK")
                            {
                                if (collMsg.Contains("The person group does not exist in the database"))
                                    sendIcs = false;
                                else
                                    sendIcs = true;

                                resRow[3] += "Error al crear la colección: " + collMsg + "\n";
                                console.WriteLine(resRow[0] + ": Error al crear la colección: " + collMsg);
                            }
                            else
                            {
                                resRow[1] = singleCollection[0][2].ToString().ToUpper();
                                resRow[2] = colDesc;
                                resRow[3] += collMsg;

                                string logMsg = $"{resRow[0]} | {resRow[1]} | {resRow[2]} | {resRow[3]}";
                                log.LogDeCambios("", "", root.BDUserCreatedBy, "Nueva Colección en CD", logMsg, "");
                                root.requestDetails += "Nueva Colección en CD: " + logMsg;
                                console.WriteLine("Nueva Colección en CD: " + logMsg);
                            }

                        }
                        catch (Exception ex)
                        {
                            sendIcs = true;
                            resRow[3] += "Error colección: " + ex.Message;
                            console.WriteLine(resRow[0] + ":  Error colección: " + ex.Message);
                        }
                    }
                    else
                    {
                        resRow[3] += "La colección ya existe";
                        console.WriteLine("La colección ya existe");
                    }

                    resDt.Rows.Add(resRow);
                }
                console.WriteLine("Enviando correo");
                if (sendIcs)
                    mail.SendHTMLMail("Falló carga de creación de colecciones <br><br>Solicitud: <br><br>" + val.ConvertDataTableToHTML(resDt), new string[] { root.BDUserCreatedBy }, root.BDClass, new string[] { "internalcustomersrvs@gbm.net" }, new string[] { filePath });
                else
                    SendResponseMail("Carga de colecciones en Control desk<br><br>", resDt, new string[] { filePath });
            }
            else
            {
                mail.SendHTMLMail("La plantilla no contiene ninguna hoja con el nombre \"Formulario Colecciones\"", new string[] { root.BDUserCreatedBy }, root.BDClass, attachments: new string[] { filePath });
            }
        }

        private void SendResponseMail(string message, DataTable resDt, string[] attach )
        {
            DataView view = new DataView(resDt);
            DataTable distinctValues = view.ToTable(true, "Supervisor");

            foreach (DataRow row in distinctValues.Rows)
            {
                List<string> sender = new List<string>
                {
                    root.BDUserCreatedBy
                };

                string supervisor = row["Supervisor"].ToString();
                DataRow[] arr;
                if (supervisor == "")
                    arr = resDt.Select("Supervisor is null");
                else 
                {
                    arr = resDt.Select("Supervisor = '" + supervisor + "'");
                    sender.Add(supervisor);
                }

                mail.SendHTMLMail(message + val.ConvertDataTableToHTML(arr.CopyToDataTable()), sender.ToArray(), root.Subject, attachments: attach);
            }
        }
    }
}
