using Microsoft.Exchange.WebServices.Data;
using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Security;
using System.Text;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;

namespace DataBotV5.Security
{
    class SecureAccess
    {
        ConsoleFormat console = new ConsoleFormat();
        ~SecureAccess()
        {

        }
        public SecureString GetPassword(int tipoMensaje)
        {
            switch (tipoMensaje)
            {
                case 1:
                    console.WriteLine(" Please type Database Access code:");
                    break;
                    //case 2:
                    //    console.WriteLine(" Please type SAP Communication User Access code:");
                    //    break;
                    //case 3:
                    //    console.WriteLine(" Please type Data Master Portal Access code:");
                    //    break;
                    //case 4:
                    //    console.WriteLine(" Please type SAP Databot Access code:");
                    //    break;
                    //case 5:
                    //    console.WriteLine(" Please type SAP Control Desk Access code:");
                    //    break;
            }
            var pwd = new SecureString();
            while (true)
            {
                ConsoleKeyInfo i = Console.ReadKey(true);
                if (i.Key == ConsoleKey.Enter)
                {
                    break;
                }
                else if (i.Key == ConsoleKey.Backspace)
                {
                    if (pwd.Length > 0)
                    {
                        pwd.RemoveAt(pwd.Length - 1);
                        Console.Write("\b \b");
                    }
                }
                else if (i.KeyChar != '\u0000') // KeyChar == '\u0000' if the key pressed does not correspond to a printable character, e.g. F1, Pause-Break, etc
                {
                    pwd.AppendChar(i.KeyChar);
                    Console.Write("*");
                }
            }
            console.WriteLine("");
            return pwd;
        }

        public string EncodePass(SecureString password)
        {
            string ps_convert = new System.Net.NetworkCredential(string.Empty, password).Password;
            string pass = string.Empty;
            byte[] encrypt = System.Text.ASCIIEncoding.ASCII.GetBytes(ps_convert);
            pass = Convert.ToBase64String(encrypt);
            return pass;
        }
        public string DecodePass(string password)
        {
            string pass = string.Empty;
            try
            {
                byte[] decrypt = Convert.FromBase64String(password);
                //pass = System.Text.Encoding.Unicode.GetString(decrypt);
                pass = System.Text.ASCIIEncoding.ASCII.GetString(decrypt);
            }
            catch (Exception) { }

            return pass;
        }


        public string RandomPassword()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(RandomString(4, true));
            builder.Append(new ValidateData().RandomNumber(1000, 9999));
            builder.Append(RandomString(2, false));
            return builder.ToString();
        }
        public string RandomString(int size, bool lowerCase)
        {
            StringBuilder builder = new StringBuilder();
            Random random = new Random();
            char ch;
            for (int i = 0; i < size; i++)
            {
                ch = Convert.ToChar(Convert.ToInt32(Math.Floor(26 * random.NextDouble() + 65)));
                builder.Append(ch);
            }
            if (lowerCase)
                return builder.ToString().ToLower();
            return builder.ToString();
        }

        public bool PrivateAccess()
        {
            string user;
            string pass = "";
            string code = "";
            bool validated = false;
            console.WriteLine(" Databot Activated");
            console.WriteLine(" Please type your ID on the space below and press ENTER");
            user = Console.ReadLine();
            for (int i = 1; i <= 4; i++)
            {
                if (user == "")
                {
                    console.WriteLine(" Please type your ID on the space below and press ENTER");
                    user = Console.ReadLine();
                }
                else
                {
                    if (user != "DMEZA" && user != "SMARIN" && user != "JEARAYA" && user != "CLGARCIA")
                    {
                        console.WriteLine(DateTime.Now + " > >  >" + " ID not recognized. Please type your ID on the space below and press ENTER");
                        user = Console.ReadLine();
                    }
                    else
                    {
                        //Si es el usuario adecuado
                        console.WriteLine(" ID seems to match according to my senses");
                        Rooting root = new Rooting();
                        root.Current_User = user;
                        console.WriteLine(" Please check your email and enter the verification code in the space below and press ENTER");
                        validated = true;
                        break;
                    }
                }
                //Los intentos exceden los 3 intentos
                if (i == 3)
                {
                    if (user != "DMEZA" && user != "SMARIN" && user != "JEARAYA" && user != "CLGARCIA")
                    {
                        console.WriteLine(" ALERT: LIMIT OF TRIES REACHED, CLOSING PROGRAM");
                        for (int z = 3; z > 0; z--)
                        {
                            console.WriteLine(z.ToString());
                            System.Threading.Thread.Sleep(1000);

                        }
                        Environment.Exit(0);
                    }
                    else
                    {
                        console.WriteLine(" ID seems to match according to my senses");
                        console.WriteLine(" Please check your email and enter the verification code in the space below and press ENTER");
                        validated = true;
                        break;
                    }

                }

            }
            if (validated == true)
            {
                pass = RandomPassword();
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                Credentials cred = new Credentials();
                service.Credentials = new WebCredentials("gbmdmbot@outlook.com", cred.password_exchange);
                // service.TraceEnabled = true;
                //  service.TraceFlags = TraceFlags.All;
                service.AutodiscoverUrl("gbmdmbot@outlook.com", RedirectionUrlValidationCallback);
                EmailMessage email = new EmailMessage(service);
                email.ToRecipients.Add(user + "@GBM.NET");
                email.Subject = "Databot Access Password";
                email.Body = new MessageBody("Code: " + pass);
                email.Send();
                #region Matar Objetos              
                System.GC.SuppressFinalize(service);
                System.GC.SuppressFinalize(email);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                #endregion
                console.WriteLine(pass);
                code = Console.ReadLine();
                if (code == pass)
                {
                    console.WriteLine(" Welcome Master " + user);
                    console.WriteLine(" Starting DataBot Services");

                }
                else
                {
                    console.WriteLine("ALERT: SECUIRITY COMPROMISED, CLOSING PROGRAM");
                    for (int z = 3; z > 0; z--)
                    {
                        console.WriteLine(z.ToString());
                        System.Threading.Thread.Sleep(1000);
                    }
                    Environment.Exit(0);
                }

            }
            return true;

        }
        public string Select_Area()
        {
            DataTable mytable = new DataTable();
            Credentials cred = new Credentials();
            string code = "";
            string valor = "";
            string sql_select = "";
            string currentvalue = "";

            string[] area_array = new String[10];
            int contador = 0;
            bool validated = false;

            console.WriteLine(" Please type one of the options below and press ENTER for start");

            #region extraer los departamentos

            try
            {
                #region Connection DB 
                sql_select = "select * from orchestrator";
                mytable = new CRUD().Select(sql_select, "databot_db");
                #endregion

                contador = 0;
                if (mytable.Rows.Count > 0)
                {
                    foreach (DataRow row in mytable.Rows)
                    {
                        valor = row[2].ToString();
                        int pos1 = Array.IndexOf(area_array, valor);
                        if (pos1 <= -1)
                        {
                            console.WriteLine("   " + row[2].ToString());
                            area_array[contador] = row[2].ToString();
                            contador++;
                        }

                    }

                }

            }
            catch (Exception ex)
            { }

            #endregion

            code = Console.ReadLine();
            for (int i = 1; i <= 4; i++)
            {
                if (code == "")
                {
                    console.WriteLine(" Please type one of the options below and press ENTER for start");

                    code = Console.ReadLine();
                }
                else
                {
                    int pos = Array.IndexOf(area_array, code);
                    if (pos > -1)
                    {
                        //Todo OK
                        console.WriteLine(" ID seems to match according to my senses");
                        validated = true;
                        break;
                    }
                    else
                    {
                        console.WriteLine(DateTime.Now + " > >  >" + " Area Type not recognized. Please type one option on the space below and press ENTER");
                        code = Console.ReadLine();
                    }
                }
                //Los intentos exceden los 3 intentos
                if (i == 3)
                {
                    int pos2 = Array.IndexOf(area_array, code);
                    if (pos2 > -1)
                    {
                        //Todo OK
                        console.WriteLine(" Code seems to match according to my senses");
                        validated = true;
                        break;
                    }
                    else
                    {
                        console.WriteLine(" ALERT: LIMIT OF TRIES REACHED, CLOSING PROGRAM");
                        for (int z = 3; z > 0; z--)
                        {
                            console.WriteLine(z.ToString());
                            System.Threading.Thread.Sleep(1000);

                        }
                        Environment.Exit(0);
                    }
                }

            }
            if (validated == true)
            {
                return code;
            }
            else
            {
                return "";
            }


        }
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            bool result = false;
            try
            {
                Uri redirectionUri = new Uri(redirectionUrl);
                if (redirectionUri.Scheme == "https")
                {
                    result = true;
                }
                return result;
            }
            catch (Exception)
            {
                return redirectionUrl.ToLower().StartsWith("https://");
            }

        }
        public void EnterLoginData()
        {
            Credentials cred = new Credentials();
            int tries = 0;
            bool correct = true;

            do
            {
                cred.pass_db1 = EncodePass(GetPassword(1));
                if (cred.ConnectDB() == true)
                {
                    correct = true;
                }
                else
                {
                    correct = false;
                }
                //cred.password_dm = EncodePass(GetPassword(3));
                //if (cred.ConnectDM() == true)
                //{
                //    correct = true;
                //}
                //else
                //{
                //    correct = false;
                //}
                //cred.password_SAPPRD = EncodePass(GetPassword(4));
                //if (cred.ConnectSAP() == true)
                //{
                //    correct = true;

                //}
                //else
                //{
                //    correct = false;
                //}
                //cred.password_CD = EncodePass(GetPassword(5));
                //if (cred.ConnectCD() == true)
                //{
                //    correct = true;
                //}
                //else
                //{
                //    correct = false;
                //}

                if (tries + 1 != 3 && correct == false)
                {
                    console.WriteLine(" Password is not correct, please try again   Tries: " + (tries + 1) + " of 3.");
                    console.WriteLine("");
                }

                tries++;

            } while (tries <= 2 && correct == false);

            if (tries == 3 && correct == false)
            {
                console.WriteLine("ALERT: SECUIRITY COMPROMISED, CLOSING PROGRAM");
                System.Threading.Thread.Sleep(4000);
                Environment.Exit(0);
            }
        }
    }
}
