using Newtonsoft.Json;
using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;

namespace DataBotV5.Logical.Projects.Contacts
{ 
    /// <summary>
    /// Clase Logical encargada de administración de contactos con SAP.
    /// </summary>

    class ContactSapSS
    {
        ValidateData val = new ValidateData();
        Credentials cred = new Credentials();
        ConsoleFormat console = new ConsoleFormat();
        Rooting root = new Rooting();
        SapVariants sap = new SapVariants();
        string mandante = "ERP";
        string mandCrm = "CRM";

        public string CreateContactSAP(ContactInfoSS contactInfo)
        {
            string resp = "";
            string fullName = $"{contactInfo.nombre} {contactInfo.apellido}, ";
            //valida la información el campo "respuesta" se llena si da error o la data esta erronea.
            contactInfo = ValData(contactInfo);
            #region SAP
            if (string.IsNullOrEmpty(contactInfo.respuesta))
            {
                console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                try
                {
                    RfcDestination destination = sap.GetDestRFC(mandCrm);
                    RfcRepository repo = destination.Repository;
                    IRfcFunction func = repo.CreateFunction("ZDM_CREATE_CONTACT");

                    #region Parametros de SAP
                    func.SetValue("CLIENTE", contactInfo.cliente);
                    func.SetValue("TRATAMIENTO", contactInfo.tratamiento);
                    func.SetValue("NOMBRE", contactInfo.nombre);
                    func.SetValue("APELLIDO", contactInfo.apellido);
                    func.SetValue("PAIS", contactInfo.pais);
                    func.SetValue("DIRECCION", contactInfo.direccion);
                    func.SetValue("CORREO", contactInfo.email);
                    IRfcTable telefonos = func.GetTable("TELEFONO");
                    telefonos.Clear();
                    foreach (phonesSS phone in contactInfo.telefonos)
                    {

                        telefonos.Append();
                        telefonos.SetValue("TELEPHONE", phone.TELEPHONE);
                        telefonos.SetValue("MOBILE", phone.MOBILE);
                        telefonos.SetValue("EXT", phone.EXT);
                    }
                    func.SetValue("IDIOMA", contactInfo.idioma);

                    func.SetValue("FUNCION", contactInfo.puesto);
                    func.SetValue("DEPARTAMENTO", contactInfo.departamento);
                    #endregion
                    #region Invocar FM
                    func.Invoke(destination);
                    #endregion

                    #region Procesar Salidas del FM
                    resp =  func.GetValue("RESPUESTA").ToString() + "<br>";
                   
                    //log de cambios base de datos
                    console.WriteLine(contactInfo.nombre + " " + contactInfo.apellido + ": " + func.GetValue("RESPUESTA").ToString());


                    #endregion
                }
                catch (Exception ex)
                {
                    resp = "Error " + contactInfo.nombre + " " + contactInfo.apellido + ": " + ex.ToString() + "<br>";
                }
            }
            else
            {
                resp = contactInfo.respuesta;
            }
            #endregion
            return resp;
        }
        public string UpdateContactSAP(ContactInfoSS contactInfo)
        {

            string resp = "";

            //validar datos
            contactInfo = ValData(contactInfo);
            if (contactInfo.contacto.All(char.IsNumber) == false || contactInfo.contacto.Substring(0, 3) != "007")
            {
                contactInfo.respuesta = "Error: ID del contacto no valido";

            }

            #region SAP
            if (!string.IsNullOrEmpty(contactInfo.respuesta))
            {
                resp = contactInfo.respuesta;
            }
            else
            {
                console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                try
                {
                    RfcDestination dest_crm = sap.GetDestRFC(mandCrm);
                    IRfcFunction modi_contact = dest_crm.Repository.CreateFunction("ZICS_BP_MODI_CONTACT");
                    modi_contact.SetValue("CLIENTE", contactInfo.cliente);
                    modi_contact.SetValue("CONTACTO", contactInfo.contacto);
                    modi_contact.SetValue("TRATAMIENTO", contactInfo.tratamiento);
                    modi_contact.SetValue("NOMBRE", contactInfo.nombre);
                    modi_contact.SetValue("APELLIDO", contactInfo.apellido);
                    modi_contact.SetValue("PAIS", contactInfo.pais);
                    modi_contact.SetValue("DIRECCION", contactInfo.direccion);
                    modi_contact.SetValue("CORREO", contactInfo.email);
                    modi_contact.SetValue("FUNCION", contactInfo.puesto);
                    modi_contact.SetValue("DEPARTAMENTO", contactInfo.departamento);
                    modi_contact.SetValue("IDIOMA", contactInfo.idioma);
                    IRfcTable telefono = modi_contact.GetTable("TELEFONO");
                    telefono.Clear();
                    foreach (phonesSS phone in contactInfo.telefonos)
                    {

                        telefono.Append();
                        telefono.SetValue("TELEPHONE", phone.TELEPHONE);
                        telefono.SetValue("MOBILE", phone.MOBILE);
                        telefono.SetValue("EXT", phone.EXT);
                    }

                    modi_contact.Invoke(dest_crm);

                    //resp = contactInfo.nombre + " " + contactInfo.apellido + ": " + func.GetValue("RESPUESTA").ToString() + "<br>";

                    #region Procesar salidas del FM y crear Json para la BD
                    DataTable ret = sap.GetDataTableFromRFCTable(modi_contact.GetTable("RET"));
                    string respuesta = modi_contact.GetValue("RETURN").ToString();
                    IRfcStructure info = modi_contact.GetStructure("CONTACT_INFO");
                    DataTable contact_info = sap.GetDataTableFromRFCStructure(modi_contact.GetStructure("CONTACT_INFO"));
                    DataTable info_tel = sap.GetDataTableFromRFCTable(info.GetTable("PHONES"));

                    string info_tel_json = JsonConvert.SerializeObject(info_tel);
                    string contact_info_json = JsonConvert.SerializeObject(contact_info);

                    contact_info_json = contact_info_json.Replace("\"PHONES\":null", "\"PHONES\":" + info_tel_json);
                    #endregion

                    if (ret.Select("TYPE = 'E'").Length > 0)
                    {
                        string ret_html = val.ConvertDataTableToHTML(ret);
                        resp = $"Error: {respuesta}<br>{ret_html}";
                    }
                    else
                    {
                        //Respuesta
                        resp = contact_info_json;
                    }
                }
                catch (Exception ex)
                {
                    resp = "Error " + contactInfo.nombre + " " + contactInfo.apellido + ": " + ex.ToString() + "<br>";
                }
            }
            #endregion
            return resp;
        }
        public ContactInfoSS ValData(ContactInfoSS contactInfo)
        {

            string fullName = $"{contactInfo.nombre} {contactInfo.apellido}, ";
            string er = $"Error: {fullName}";

            if (contactInfo.cliente.All(char.IsNumber) == false || contactInfo.cliente.Substring(0, 3) != "001" && contactInfo.cliente.Substring(0, 1) != "1")
            {
                contactInfo.respuesta = er + "ID del cliente no valido";

            }

            contactInfo.cliente = contactInfo.cliente.PadLeft(10, '0');
            //significa que es la plantilla de Datos Maestros 01. Sra / 02. Sr
            if (contactInfo.tratamiento.Contains('.'))
            {
                contactInfo.tratamiento = contactInfo.tratamiento.Substring(0, 2);
            }

            contactInfo.tratamiento = contactInfo.tratamiento.PadLeft(4, '0');

            contactInfo.nombre = val.RemoveSpecialChars(contactInfo.nombre, 1);
            contactInfo.nombre = val.RemoveChars(contactInfo.nombre);
            contactInfo.apellido = val.RemoveSpecialChars(contactInfo.apellido, 1);
            contactInfo.apellido = val.RemoveChars(contactInfo.apellido);

            contactInfo.pais = contactInfo.pais.Substring(0, 2);

            contactInfo.direccion = val.RemoveSpecialChars(contactInfo.direccion, 1);
            if (contactInfo.direccion.Length > 60)
            { contactInfo.direccion = contactInfo.direccion.Substring(0, 60); }
            contactInfo.direccion = val.RemoveChars(contactInfo.direccion);

            contactInfo.email = contactInfo.email.ToLower().Trim();
            contactInfo.email = val.RemoveChars(contactInfo.email);
            contactInfo.email = val.RemoveEnne(contactInfo.email);
            if ((contactInfo.email.IndexOf("@") + 1) == 0 || (contactInfo.email.IndexOf("ñ") + 1) > 0 || !val.ValidateEmail(contactInfo.email))
            {
                contactInfo.respuesta = er + $"{contactInfo.email} Por favor ingresar un formato correcto de email";
            }
            List<phonesSS> telefonos = new List<phonesSS>();
            foreach (phonesSS phone in contactInfo.telefonos)
            {
                if (phone.TELEPHONE.Substring(0, 1) == "(")
                { phone.TELEPHONE = phone.TELEPHONE.Substring(5, phone.TELEPHONE.Length - 5); }

                if ((phone.TELEPHONE.IndexOf("ext") + 1) > 0)
                { phone.TELEPHONE = phone.TELEPHONE.Substring(0, phone.TELEPHONE.IndexOf("ext") - 1); }

                phone.TELEPHONE = phone.TELEPHONE.Replace("+", "");

                phone.TELEPHONE = val.EditPhone(phone.TELEPHONE);
                phone.MOBILE = val.EditPhone(phone.MOBILE);
                if (phone.EXT == null)
                {
                    phone.EXT = "";
                }

                telefonos.Add(phone);
            }
            contactInfo.telefonos = telefonos;

            //en caso de creacion DM llega con 2 digitos

            //en caso de actualizacion o creacion masiva por S&S llega con los 4 caracteres
            if (!string.IsNullOrWhiteSpace(contactInfo.puesto))
            {
                
                if (contactInfo.puesto.Contains("-"))
                {
                    //en caso de creacion masiva DM llega como 0001 - puesto por lo que se corta para que quede como 2 o 1 digito
                    contactInfo.puesto = contactInfo.puesto.Substring(0, 4).ToUpper().TrimStart('0');

                }

            }
            if (!string.IsNullOrWhiteSpace(contactInfo.departamento))
            {
                if (contactInfo.departamento.Contains("-"))
                {
                    contactInfo.departamento = contactInfo.departamento.Substring(0, 4).ToUpper().TrimStart('0');

                }

            }
            return contactInfo;
        }
    }
    class ContactInfoSS
    {
        public string respuesta { get; set; } //optional
        public string cliente { get; set; }
        public string contacto { get; set; } //optional
        public string tratamiento { get; set; }
        public string nombre { get; set; }
        public string apellido { get; set; }
        public string pais { get; set; }
        public string direccion { get; set; }
        public string email { get; set; }
        public List<phonesSS> telefonos { get; set; }
        public string idioma { get; set; }
        public string puesto { get; set; } //optional
        public string departamento { get; set; } //optional
    }
    class phonesSS
    {
        public string TELEPHONE { get; set; }
        public string EXT { get; set; }
        public string MOBILE { get; set; }
    }
}
