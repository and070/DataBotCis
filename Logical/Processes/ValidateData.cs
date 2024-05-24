using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Data;
using System.Globalization;
using DataBotV5.Data.Database;
using DataBotV5.Automation.WEB.H2HCredomatic;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using System.Data.SqlClient;
using DataBotV5.App.Global;
using DataBotV5.Data.SAP;
using PhoneNumbers;
using System.Linq;

namespace DataBotV5.Logical.Processes
{
    /// <summary>
    /// Clase Logical encargada de validar data.
    /// </summary>
    class ValidateData
    {
        #region Variables Globales
        readonly CRUD crud = new CRUD();

        #endregion

        /// <summary>
        /// Método para validar si un e-mail es correcto.
        /// </summary>
        /// <param name="email"></param>
        /// <returns>Devuelve un bool si el método es correcto.</returns>
        public bool ValidateEmail(string email)
        {
            string expression = "\\w+([-+.']\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*";

            if (Regex.IsMatch(email, expression))
            {
                if (Regex.Replace(email, expression, string.Empty).Length == 0)
                {
                    return true;
                }
            }
            return false;
        }
        /// <summary>
        /// Método para corregir el número telefónico.
        /// </summary>
        /// <param name="phone"></param>
        /// <returns>Devuelve un String con el teléfono corregido.</returns>
        public string EditPhone(string phone)
        {
            try
            {

                if (phone != "")
                {
                    int tel;
                    phone = phone.ToUpper();
                    if (phone.Length > 30 | phone.Contains("tel") | phone.Contains("/") | phone.Contains(","))
                    {
                        phone = "";
                    }

                    if (phone.Substring(0, 1) == "(")
                    { phone = phone.Substring(5, phone.Length - 5); }

                    tel = phone.IndexOf("EXT");
                    if (tel > 0)
                    {
                        phone = phone.Substring(0, tel - 1);
                    }
                    phone = phone.Replace("-", "");
                    phone = phone.Replace("+", "");
                    phone = phone.Trim();
                }

            }
            catch (Exception)
            {
                phone = "";
            }
            return phone;
        }
        /// <summary>
        /// Verificar si el telefono concuerda con el formato del país
        /// </summary>
        /// <param name="country">CR, PA, NI, GT, CO, US, HN, SV, DO</param>
        /// <returns></returns>
        public int GetExpectedPhoneDigitsByCountry(string country)
        {
            PhoneNumberUtil phoneNumberUtil = PhoneNumberUtil.GetInstance();

            try
            {
                PhoneNumber exampleNumber = phoneNumberUtil.GetExampleNumberForType(country, PhoneNumberType.MOBILE);

                if (exampleNumber != null)
                {
                    int phoneNumberLength = phoneNumberUtil.GetNationalSignificantNumber(exampleNumber).Length;
                    Console.WriteLine($"Expected phone number length for {country}: {phoneNumberLength}");
                    return phoneNumberLength;
                }
                else
                {
                    Console.WriteLine($"Could not retrieve an example phone number for {country}");
                    return 0;
                }
            }
            catch (NumberParseException e)
            {
                Console.WriteLine($"Error parsing phone number: {e.Message}");
                return 0;
            }
        }

        /// <summary>
        /// Método que elimina tildes.
        /// </summary>
        /// <param name="text"></param>
        /// <returns>Devuelve un String sin tildes.</returns>
        public string RemoveAccents(string text)
        {

            text = RemoveSpecialChars(text, 1);
            return text;
        }
        /// <summary>
        /// Método que quita los carcateres especiales.
        /// </summary>
        /// <param name="text">String de entrada</param>
        /// <param name="type">1 no quita puntos ni comas, 2 quita todo</param>
        /// <returns></returns>
        public string RemoveSpecialChars(string text, int type)
        {
            string resultado = "";
            //Elimina caracteres especiales
            Regex reg;
            if (type == 1)
            {
                reg = new Regex("[*'\"&+#^><]"); //•
            }
            else
            {
                reg = new Regex("[*'\",&+#.^><]");
            }

            text = reg.Replace(text, string.Empty);
            //Convierte el texto bytes y destruye las mayusculas
            byte[] bytesTemporales;
            bytesTemporales = System.Text.Encoding.GetEncoding("ISO-8859-8").GetBytes(text);
            text = System.Text.Encoding.UTF8.GetString(bytesTemporales);
            if (type == 2)
            {
                resultado = text;
                resultado = resultado.Replace("-", "");
            }
            else
            {
                resultado = text;
            }
            return resultado;
        }
        /// <summary>
        /// Método para eliminar la 'ñ'.
        /// </summary>
        /// <param name="text"></param>
        /// <returns>Devuelve un String sin el caracter 'ñ'.</returns>
        public string RemoveEnne(string text)
        {
            string resultado = "";
            resultado = text.Replace("Ñ", "N");
            resultado = text.Replace("ñ", "n");
            return resultado;
        }
        /// <summary>
        /// Método para validar SB Form.
        /// </summary>
        /// <param name="pais_ibm"></param>
        /// <param name="useopp"></param>
        /// <param name="oppo"></param>
        /// <param name="usespecial"></param>
        /// <param name="prevbid"></param>
        /// <param name="priceupdate"></param>
        /// <param name="priceupjusti"></param>
        /// <param name="customer"></param>
        /// <param name="brand"></param>
        /// <param name="sedojusti"></param>
        /// <param name="soleprocur"></param>
        /// <param name="bpjusti"></param>
        /// <param name="swma"></param>
        /// <param name="renew"></param>
        /// <param name="totalprice"></param>
        /// <param name="customerprice"></param>
        /// <returns> Retorna un bool determinando si es válido el SB Form.</returns>
        public bool ValidateSBForm(string pais_ibm, string useopp, string oppo, string usespecial,
                                       string prevbid, string priceupdate, string priceupjusti, string customer,
                                       string brand, string sedojusti, string soleprocur, string bpjusti,
                                             string swma, string renew, string totalprice, string customerprice)
        {

            bool respuesta;

            if (pais_ibm == "" || useopp == "" || customer == "" || usespecial == "" || brand == "" || sedojusti == "" || bpjusti == "" || soleprocur == "")
            {
                respuesta = false;
            }
            else
            {
                respuesta = true;
            }

            if (useopp == "Yes" && oppo == "")
            {
                respuesta = false;
            }
            else
            {
                respuesta = true;
            }

            if (priceupdate == "Yes" && priceupjusti == "")
            {
                respuesta = false;
            }
            else
            {
                respuesta = true;
            }

            if (usespecial == "Yes" && prevbid == "")
            {
                respuesta = false;
            }
            else
            {
                respuesta = true;
            }

            if (respuesta == true)
            { return true; }
            else
            { return false; }

        }
        /// <summary>
        /// Método que devuelve un mensaje detallando un error de alguna parte del catch de código.
        /// </summary>
        /// <param name="message">Se detalla el mensaje.</param>
        /// <param name="detail">Campo para el detalle.</param>
        /// <param name="file">Campo para el archivo.</param>
        /// <param name="management">Campo para la gestión.</param>
        /// <param name="line">Campo para la línea.</param>
        /// <returns>Devuelve un String con hora, mensaje, error, detalle del error, archivo que genera error y la gestión del error.</returns>
        public string LogErrors(string message, string detail, string file, string management, int line)
        {
            string retorno;

            retorno = "----------------------------------------------------------------" + "\r\n";
            retorno = retorno + "     Hora del error: " + DateTime.Today + "\r\n";
            retorno = retorno + "------------------------------------------------------" + "\r\n";
            retorno = retorno + "Mensaje del error: " + "\r\n" + "\r\n";
            retorno = retorno + message + "\r\n" + "\r\n";
            retorno = retorno + "Mensaje del error detallado: " + "\r\n" + "\r\n";
            retorno = retorno + detail + "\r\n" + "\r\n";
            retorno = retorno + "------------------------------------------------------" + "\r\n";
            retorno = retorno + "     Archivo que genera el error: " + file + "\r\n";
            retorno = retorno + "     Gestion que genera el error: " + management + "\r\n";
            if (line > 0)
            {
                retorno = retorno + "     Linea de archivo que genera el error: " + line + "\r\n";
            }
            retorno = retorno + "------------------------------------------------------" + "\r\n";

            return retorno;
        }
        /// <summary>
        /// Método para averiguar el número de la semana del mes.
        /// </summary>
        /// <param name="date"></param>
        /// <returns> Devuelve un número tipo int con el número de la semana.</returns>
        public int GetWeekNumberOfMonth(DateTime date)
        {
            date = date.Date;
            DateTime firstMonthDay = new DateTime(date.Year, date.Month, 1);
            DateTime firstMonthMonday = firstMonthDay.AddDays((DayOfWeek.Monday + 7 - firstMonthDay.DayOfWeek) % 7);
            if (firstMonthMonday > date)
            {
                firstMonthDay = firstMonthDay.AddMonths(-1);
                firstMonthMonday = firstMonthDay.AddDays((DayOfWeek.Monday + 7 - firstMonthDay.DayOfWeek) % 7);
            }
            return (date - firstMonthMonday).Days / 7 + 1;
        }
        /// <summary>
        /// Método para determinar el mes según la fecha suministrada.
        /// </summary>
        /// <param name="date"></param>
        /// <returns>Devuelve un String con el mes.</returns>
        public string Month(DateTime date)
        {
            date = date.Date;
            DateTime firstMonthDay = new DateTime(date.Year, date.Month, 1);
            DateTime firstMonthMonday = firstMonthDay.AddDays((DayOfWeek.Monday + 7 - firstMonthDay.DayOfWeek) % 7);
            if (firstMonthMonday > date)
            {
                firstMonthDay = firstMonthDay.AddMonths(-1);
                firstMonthMonday = firstMonthDay.AddDays((DayOfWeek.Monday + 7 - firstMonthDay.DayOfWeek) % 7);
            }
            return firstMonthDay.ToString("MMMM", new CultureInfo("es-ES"));
        }
        /// <summary>
        /// Método para determinar el ShipPoint del país segun su nombre.
        /// </summary>
        /// <param name="country"></param>
        /// <returns>Devuelve un String con el homólogo del país.</returns>
        public bool HomologousCountry(string country)
        {
            string[] invalidCountries = { "CR01", "CO01", "DO01", "ES01", "GU01", "HO01", "MD01", "NI01", "PA01", "BV01", "ITC1", "LCFL", "LCVE", "WTC1" };

            return invalidCountries.Contains(country);
        }
        /// <summary>
        /// Método para convertir un DataTable a HTML.
        /// </summary>
        /// <param name="dt"></param>
        /// <returns>Retorna un String con el HTML.</returns>
        public string ConvertDataTableToHTML(DataTable dt)
        {
            string html = "<table border=\"1\">";
            //add header row
            html += "<tr>";
            for (int i = 0; i < dt.Columns.Count; i++)
                html += "<td>" + dt.Columns[i].ColumnName + "</td>";
            html += "</tr>";
            //add rows
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                html += "<tr>";
                for (int j = 0; j < dt.Columns.Count; j++)
                    html += "<td>" + dt.Rows[i][j].ToString() + "</td>";
                html += "</tr>";
            }
            html += "</table>";
            return html;
        }
        /// <summary>
        /// Método para convertir un DataTable a Markdown.
        /// </summary>
        /// <param name="dt"></param>
        /// <returns>Retorna un String con el Markdown.</returns>
        public string ConvertDataTableToMarkdown(DataTable dt)
        {
            StringBuilder sb = new StringBuilder();

            foreach (DataColumn column in dt.Columns)
                sb.Append("|").Append(column.ColumnName.Trim());

            sb.AppendLine("|");

            for (int i = 0; i < dt.Columns.Count; i++)
                sb.Append("|----");

            sb.AppendLine("|");

            foreach (DataRow row in dt.Rows)
            {
                foreach (object val in row.ItemArray)
                    sb.Append("|").Append(Convert.ToString(val).Trim());

                sb.AppendLine("|");
            }

            return sb.ToString();
        }
        /// <summary>
        /// Método para eliminar chars de una cadena de texto incluyendo un carácteres especiales como $, USD, *, ;, &.  .
        /// </summary>
        /// <param name="valor"></param>
        /// <returns>Devuelve un String sin chars ni caracteres especiales.</returns>
        /// <summary>
        /// Método para eliminar chars únicamente. Elimina las tildes del todo
        /// </summary>
        /// <param name="text"></param>
        /// <returns>Devuelve un String sin cahars</returns>
        public string RemoveChars(string text)
        {
            text = Regex.Replace(text, @"[^\u0000-\u007F]", "").Trim();
            text = text.Replace("\t", "");
            text = text.Replace("\"", "");
            return text;
        }
        /// <summary>
        /// Método para quitar chars además que quita carácteres especiales.
        /// </summary>
        /// <param name="value"></param>
        /// <returns>Devuelce un String limpio.</returns>
        public string Clean(string value)
        {
            value = value.Replace("á", "a"); value = value.Replace("é", "e"); value = value.Replace("í", "i"); value = value.Replace("ó", "o"); value = value.Replace("ú", "u"); value = value.Replace("ñ", "n");
            value = value.Replace("Á", "A"); value = value.Replace("É", "E"); value = value.Replace("Í", "I"); value = value.Replace("Ó", "O"); value = value.Replace("Ú", "U"); value = value.Replace("Ñ", "N");
            value = value.Replace("\n", "");
            value = value.Replace("\r", "");
            value = RemoveSpecialChars(value, 1);

            return value;
        }
        /// <summary>
        /// Método para determinar si el string es un número.
        /// </summary>
        /// <param name="text"></param>
        /// <returns>Devuelve un bool en caso si es un número.</returns>
        public bool IsNum(string text)
        {
            foreach (char c in text)
            {
                if (c < '0' || c > '9')
                    return false;
            }
            return true;
        }
        /// <summary>
        /// Método para generar un número random.
        /// </summary>
        /// <param name="min"></param>
        /// <param name="max"></param>
        /// <returns>Devuelve un número tipo Int.</returns>
        public int RandomNumber(int min, int max)
        {
            Random random = new Random();
            return random.Next(min, max);
        }
        /// <summary>Método para convertir un campo con su separador.</summary>
        public string ExtractFieldWithSeparator(string separator, string inputField)
        {
            //0001 - sr
            //si no lo encuentra el index da -1 entonces ud le suma 1 para que sea 0
            //si largo = 0 entonces no encontro el separador

            int separatorIndex = inputField.IndexOf(separator) + 1;
            if (separatorIndex == 0)
                separatorIndex = inputField.Length + 1;

            string extractedField = inputField.Substring(0, separatorIndex - 1).Trim();

            return extractedField;
        }
        /// <summary>
        /// Método para obtener el teléfono en formato E164
        /// </summary>
        /// <param name="phoneNumber">El número de teléfono con cualquier formato(Debe tener el +)</param>
        /// <returns>el teléfono en formato E164</returns>
        public string ParsePhoneNumberToE164(string phoneNumber)
        {
            string res;

            try
            {
                PhoneNumberUtil phoneNumberUtil = PhoneNumberUtil.GetInstance();
                PhoneNumber parsedPhoneNumber = phoneNumberUtil.Parse(phoneNumber, "ZZ");
                string region = phoneNumberUtil.GetRegionCodeForNumber(parsedPhoneNumber);
                parsedPhoneNumber = phoneNumberUtil.Parse(phoneNumber, region);
                res = phoneNumberUtil.Format(parsedPhoneNumber, PhoneNumberFormat.E164);
            }
            catch (Exception ex)
            {
                res = "ERROR: " + ex.Message;
            }

            return res;
        }
        /// <summary>
        /// Método para obtener el Id de empleado de SAP de un user de AD
        /// </summary>
        /// <param name="email">El email del Usuario SIN @gbm.net</param>
        /// <returns></returns>
        public string GetEmployeeID(string email)
        {
            SapVariants sap = new SapVariants();
            string employeeID;
            DataTable idTable = crud.Select("SELECT UserID FROM digital_sign WHERE email = '" + email + "@GBM.NET'", "MIS");
            if (idTable.Rows.Count > 0)
                employeeID = idTable.Rows[0][0].ToString();
            else
            {
                Dictionary<string, string> parameters = new Dictionary<string, string> { ["USUARIO"] = email };
                IRfcFunction func = sap.ExecuteRFC("ERP", "ZFD_GET_USER_DETAILS", parameters);
                employeeID = func.GetValue("IDCOLABC").ToString();
            }

            return employeeID;
        }


        #region Métodos de apoyo para SAP
        /// <summary>
        /// Método que retorna  el código desde SAP por mandante.
        /// </summary>
        /// <param name="country"></param>
        /// <returns>Retorna un String con el código.</returns>
        public string CocodeSap(string account, string mandante)
        {
            string cocode = "";
            try
            {
                Dictionary<string, string> parameters = new Dictionary<string, string>();
                parameters["BANK_ACCOUNT"] = account;

                IRfcFunction func = new SapVariants().ExecuteRFC(mandante, "ZFI_GET_CCODE_BNKACCOUNT", parameters);

                cocode = func.GetValue("COMPANY_CODE").ToString();
            }
            catch (Exception)
            { cocode = ""; }


            return cocode;
        }
        /// <summary>
        /// Método que retorna un el código por país.
        /// </summary>
        /// <param name="country"></param>
        /// <returns>Retorna un String con el código.</returns>
        public string Cocode(string country)
        {
            string cocode = "";
            switch (country)
            {
                case "SV":
                    cocode = "GBSV";
                    break;
                case "PA":
                    cocode = "GBPA";
                    break;
                case "CR":
                    cocode = "GBCR";
                    break;
                case "DO":
                    cocode = "GBDR";
                    break;
                case "HN":
                    cocode = "GBHN";
                    break;
                case "NI":
                    cocode = "GBNI";
                    break;
                case "US":
                    cocode = "GBMD";
                    break;
                case "CO":
                    cocode = "GBCO";
                    break;
                case "VE":
                    cocode = "LCVE";
                    break;
                default:
                    cocode = "";
                    break;

            }

            return cocode;
        }
        /// <summary>
        /// Método que retorna el owner de SQL.
        /// </summary>
        /// <param name="account"></param>
        /// <returns>Retorna un String con el owner.</returns>
        public string OwnerSql2(string account)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string owner = "";
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB 
                sql_select = "select * from cuentas_bac where cuenta = " + account;
                //mytable = crud.Select("Databot", sql_select, "finanzas");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    owner = mytable.Rows[0][2].ToString();
                }

            }
            catch (Exception)
            {
                owner = "";
            }

            return owner;
        }
        /// <summary>
        /// Método que cuenta el owner.
        /// </summary>
        /// <returns> Retorna un Dictionary</returns>
        public Dictionary<string, string> OwnerCont()
        {
            Dictionary<string, string> cuentas = new Dictionary<string, string>();
            DataTable mytable = new DataTable();
            try
            {
                string sql_select = "SELECT country, accountant FROM `countryAccountant`";
                mytable = crud.Select(sql_select, "h2h_finance_db");

                if (mytable.Rows.Count > 0)
                {
                    for (int i = 0; i < mytable.Rows.Count; i++)
                    {
                        string pais = mytable.Rows[i][0].ToString();
                        string contador = mytable.Rows[i][1].ToString();
                        cuentas[pais] = contador;
                    }
                }
            }
            catch (Exception ex)
            {
                new ConsoleFormat().WriteLine(ex.ToString());

            }
            return cuentas;
        }
        /// <summary>
        /// Método que retorna una lista con cuentas owner sql.
        /// </summary>
        /// <returns>Retorna una lista.</returns>
        public List<accounts> OwnerSql()
        {
            List<accounts> cuentas = new List<accounts>();
            DataTable mytable = new DataTable();
            try
            {
                string sql_select = "SELECT * FROM `bacAccount`";
                mytable = crud.Select(sql_select, "h2h_finance_db");

                if (mytable.Rows.Count > 0)
                {
                    for (int i = 0; i < mytable.Rows.Count; i++)
                    {
                        accounts table = new accounts();
                        table.account = mytable.Rows[i]["account"].ToString();
                        table.owner = mytable.Rows[i]["owner"].ToString();
                        table.active = mytable.Rows[i]["active"].ToString();
                        cuentas.Add(table);

                    }
                }
            }
            catch (Exception ex)
            {
                new ConsoleFormat().WriteLine(ex.ToString());

            }
            return cuentas;
        }

        #endregion

    }
}
