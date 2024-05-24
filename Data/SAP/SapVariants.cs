using System;
using System.Data;
using SAP.Middleware.Connector;
using DataBotV5.Data.Root;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using System.Collections.Generic;
using DataBotV5.Data.Database;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using DataBotV5.App.ConsoleApp;

namespace DataBotV5.Data.SAP
{
    /// <summary>
    /// Clase Data con todas las variantes de SAP.
    /// </summary>
    class SapVariants : IDisposable
    {

        Credentials.Credentials cred = new Credentials.Credentials();
        ConsoleFormat console = new ConsoleFormat();
        Rooting root = new Rooting();
        ProcessInteraction proc = new ProcessInteraction();
        private bool disposedValue;

        public bool Enable32BitAppOnWin64 { get; set; }
        public static SAPFEWSELib.GuiApplication GuiApp { get; set; }
        public static SAPFEWSELib.GuiApplication GuiApp_qa { get; set; }
        public static SAPFEWSELib.GuiConnection connection { get; set; }
        public static SAPFEWSELib.GuiConnection connection_qa { get; set; }
        public static SAPFEWSELib.GuiSession session { get; set; }
        public static SAPFEWSELib.GuiSession session_qa { get; set; }
        public static SAPFEWSELib.GuiFrameWindow frame { get; set; }
        public static SAPFEWSELib.GuiFrameWindow frame_qa { get; set; }

        /// <summary>
        /// se logea a SAP Logon Gui (en frontend como una macro)
        /// </summary>
        /// <param name="System">ERP O CRM</param>
        /// <param name="mandante">120, 260, 300</param>
        public void LogSAP(string system, [Optional] int mandante)
        {
            #region default
            mandante = checkDefault(system, mandante);
            if (mandante == 0)
            {
                return;
            }
            #endregion
            string mand;
            string usuario = "";
            string contra = "";

            if (mandante == 300)
            {
                mand = "ERP-PRD";
                usuario = cred.username_SAPPRD;
                contra = cred.password_SAPPRD;
            }
            else if (mandante == 500)
            {
                mand = "CRM-PRD";
                usuario = cred.username_SAPPRD;
                contra = cred.password_SAPPRD;
            }
            else if (mandante == 120)
            {
                mand = "ERP 120 DEV";
                usuario = "RPAUSER";
                contra = "GbmMIS2019$*";
            }
            else if (mandante == 110)
            {
                mand = "ERP 110 DEV";
                usuario = "RPAUSER";
                contra = "GbmMIS2020$*";
            }
            else if (mandante == 260)
            {
                mand = "ERP-QA";
                usuario = "RPAUSER";
                contra = cred.password_QA_ERP;
            }
            else if (mandante == 460)
            {
                mand = "CRM-QA";
                usuario = "RPAUSER";
                contra = cred.password_QA_CRM;
            }
            else if (mandante == 420)
            {
                mand = "CRM-DEV";
                usuario = "";
                contra = "";
            }
            else
            {
                mand = "ERP 120 DEV";
                usuario = "RPAUSER";
                contra = "GbmMIS2019$*";
            }

            System.Diagnostics.Process.Start(@"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe");
            System.Threading.Thread.Sleep(5000);


            try
            {
                Enable32BitAppOnWin64 = true;


                SapROTWr.CSapROTWrapper sapROTWrapper = new SapROTWr.CSapROTWrapper();

                //Get the ROT Entry for the SAP Gui to connect to the COM
                object SapGuilRot = sapROTWrapper.GetROTEntry("SAPGUI");
                //Get the reference to the Scripting Engine
                object engine = SapGuilRot.GetType().InvokeMember("GetScriptingEngine", System.Reflection.BindingFlags.InvokeMethod, null, SapGuilRot, null);
                //Get the reference to the running SAP Application Window
                SapVariants.GuiApp = (SAPFEWSELib.GuiApplication)engine;
                //Get the reference to the first open connection
                //GuiConnection connection = (GuiConnection)GuiApp.Connections.ElementAt(mand);
                SapVariants.connection = (SAPFEWSELib.GuiConnection)GuiApp.OpenConnection(mand);
                //get the first available session
                SapVariants.session = (SAPFEWSELib.GuiSession)connection.Children.ElementAt(0);
                //Get the reference to the main "Frame" in which to send virtual key commands
                SapVariants.frame = (SAPFEWSELib.GuiFrameWindow)session.FindById("wnd[0]");
                //frame.Iconify();

                ((SAPFEWSELib.GuiTextField)session.FindById("wnd[0]/usr/txtRSYST-MANDT")).Text = mandante.ToString();

                ((SAPFEWSELib.GuiTextField)session.FindById("wnd[0]/usr/txtRSYST-BNAME")).Text = usuario;  //"dmeza";

                ((SAPFEWSELib.GuiTextField)session.FindById("wnd[0]/usr/pwdRSYST-BCODE")).Text = contra; // "Dominogbm\"0";

                ((SAPFEWSELib.GuiTextField)session.FindById("wnd[0]/usr/txtRSYST-LANGU")).Text = "EN";

                frame.SendVKey(0);
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message.ToString());

            }





        }

        /// <summary>
        /// Mata todas las sesiones abiertas de SAP Gui logon
        /// </summary>
        public void KillSAP()
        {
            try
            {
                SapVariants.connection.CloseSession("ses[0]");
                SapVariants.connection.CloseConnection();
                SapVariants.session = null;
                SapVariants.connection = null;
                SapVariants.GuiApp = null;
                proc.KillProcess("saplogon", false);
            }
            catch (Exception)
            { }

        }

        /// <summary>
        /// Revisa si el usuario de SAP del robot esta siendo usado, se valida en PRD porque solo existe 1 usuario
        /// </summary>
        /// <param name="mandante"></param>
        /// <returns></returns>
        public bool CheckLogin(string system, [Optional] int mandante)
        {
            #region default
            mandante = checkDefault(system, mandante);
            if (mandante == 0)
            {
                return true;
            }
            #endregion
            bool activo = false;
            DataTable mytable = new DataTable();
            Credentials.Credentials cred = new Credentials.Credentials();
            string select;

            select = "Select active from activeSap where client = " + mandante;
            mytable = new CRUD().Select(select, "databot_db");

            if (mytable.Rows.Count > 0)
            {
                // Console.WriteLine(mytable.Rows[0][0]);
                activo = Convert.ToBoolean(mytable.Rows[0][0]);
            }

            return activo;
        }

        /// <summary>
        /// Método para bloquear o desbloquear el RPAUSER - mandante de SAP, ya que hay un solo usuario de SAP para 
        /// las 5 instancias de las máquinas virtuales, y no se pueden usar de manera simultanea, por tanto se "bloquea" para que las demás
        /// instancias no intenten usarlo. 
        /// Mandante: variable tipo int con el mandante de SAP a bloquear.
        /// Block: variable tipo int. 1 para bloquear, 0 para desbloquear.
        /// </summary>
        /// <param name="mandante"> es el mandante que quiere bloquear o desbloquear: 120, 260, 300</param>
        /// <param name="block">0 para desbloquear, 1 para bloquear</param>
        public void BlockUser(string system, int block, [Optional] int mandante)
        {
            #region default
            mandante = checkDefault(system, mandante);
            if (mandante == 0)
            {
                return;
            }
            #endregion

            string sql_update = $"UPDATE activeSap SET `active` = {block}, class = '{root.BDClass}'  where client = " + mandante;
            new CRUD().Update(sql_update, "databot_db");

            if (block == 1)
                console.WriteLine("User de sap bloqueado");
            else
                console.WriteLine("User de sap desbloqueado");
        }

        public string AmEmail(string AM)
        {
            string email = "";
            try
            {
                console.WriteLine(" Conectado con SAP ERP");
                Dictionary<string, string> parameters = new Dictionary<string, string>();
                parameters["BP"] = AM;

                IRfcFunction func = ExecuteRFC("ERP", "ZDM_READ_BP", parameters);
                email = func.GetValue("EMAIL").ToString();

            }
            catch (Exception)
            {
                email = "dmeza@gbm.net";
            }

            return email;
        }

        public DataTable GetSapTable(string tabla, string system, [Optional] int mandante)
        {
            DataTable table = new DataTable();
            try
            {
                console.WriteLine(" Conectado con SAP ERP");
                Dictionary<string, string> parameters = new Dictionary<string, string>();
                parameters["QUERY_TABLE"] = tabla;

                IRfcFunction func = ExecuteRFC(system, "RFC_READ_TABLE", parameters, mandante);
                table = GetDataTableFromRFCTable(func.GetTable("DATA"));



            }
            catch (Exception ex)
            {

            }

            return table;
        }
        public DataTable GetDataTableFromRFCTable(IRfcTable lrfcTable)
        {
            //sapnco_util
            DataTable loTable = new DataTable();

            //... Create ADO.Net table.
            for (int liElement = 0; liElement < lrfcTable.ElementCount; liElement++)
            {
                RfcElementMetadata metadata = lrfcTable.GetElementMetadata(liElement);
                if (metadata.DataType.ToString() == "TABLE")
                    loTable.Columns.Add(metadata.Name, typeof(DataTable));
                else
                    loTable.Columns.Add(metadata.Name);
            }

            //... Transfer rows from lrfcTable to ADO.Net table.
            foreach (IRfcStructure row in lrfcTable)
            {
                DataRow ldr = loTable.NewRow();
                for (int liElement = 0; liElement < lrfcTable.ElementCount; liElement++)
                {
                    RfcElementMetadata metadata = lrfcTable.GetElementMetadata(liElement);
                    try { ldr[metadata.Name] = row.GetString(metadata.Name); }
                    catch (Exception)
                    {
                        //ldr[metadata.Name] = "Es otra Tabla, por favor tomarla aparte";
                    }
                }
                loTable.Rows.Add(ldr);
            }

            loTable.TableName = lrfcTable.Metadata.Name;

            return loTable;
        }
        public DataTable GetDataTableFromRFCStructure(IRfcStructure IRfcStructure)
        {
            //sapnco_util
            DataTable loTable = new DataTable();

            //... Create ADO.Net table.
            for (int liElement = 0; liElement < IRfcStructure.ElementCount; liElement++)
            {
                RfcElementMetadata metadata = IRfcStructure.GetElementMetadata(liElement);
                if (metadata.DataType.ToString() == "TABLE")
                    loTable.Columns.Add(metadata.Name, typeof(DataTable));
                else
                    loTable.Columns.Add(metadata.Name);
            }

            ////... Transfer rows from lrfcTable to ADO.Net table.
            DataRow ldr = loTable.NewRow();
            for (int liElement = 0; liElement < IRfcStructure.ElementCount; liElement++)
            {
                RfcElementMetadata metadata = IRfcStructure.GetElementMetadata(liElement);
                try { ldr[metadata.Name] = IRfcStructure.GetString(metadata.Name); }
                catch (Exception)
                {
                    //ldr[metadata.Name] = "Es otra Tabla, por favor tomarla aparte";
                }
            }
            loTable.Rows.Add(ldr);

            return loTable;
        }

        /// <summary>
        /// Método para ejecutar una function module de SAP, el cual através de parámetros (Dictionary) puede invocar un FM y corre un webservice.
        /// </summary>
        /// <param name="system">ERP o CRM (dependiendo de la variable de mandante por defecto toma dev, qa, prd.</param>
        /// <param name="mandante">opcional El número mandante de SAP: 120 (ERP-DEV), 420(CRM-DEV), 260(ERP-QA), 460(CRM-QA), 300(ERP-PRD), 500(CRM-PRD).</param>
        /// <param name="FunctionModule">Function module SAP: se puede encontrar en la transacción de SAP SE37. </param>
        /// <param name="parametros">Es un diccionario donde la llave es el input de la function module y el valor a llevar.</param>
        /// <returns>Retorna IRfcFunction.</returns>
        public IRfcFunction ExecuteRFC(string system, string FunctionModule, Dictionary<string, string> parametros, [Optional] int mandante)
        {
            IRfcFunction func1 = null;
            #region default
            mandante = checkDefault(system, mandante);
            if (mandante == 0)
            {
                return func1;
            }
            #endregion
            cred.IngresarAmbiente(mandante);
            RfcDestination destination1 = RfcDestinationManager.GetDestination(cred.parametros);

            RfcRepository repo1 = destination1.Repository;
            func1 = repo1.CreateFunction(FunctionModule);

            #region Parametros de SAP
            foreach (KeyValuePair<string, string> pair in parametros)
            {
                string campo = pair.Key.ToString();
                string valor = pair.Value.ToString();
                func1.SetValue(campo, valor);
            }
            #endregion

            #region Invocar FM
            func1.Invoke(destination1); //corre el ws
            #endregion

            return func1;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="zys"></param>
        /// <param name="mand"></param>
        /// <returns></returns>
        public int checkDefault(string zys, int mand)
        {
            if (mand == 0) //significa no mando el mandate
            {
                string amb = Start.enviroment;
                if (amb == "PRD")
                {
                    mand = (zys == "CRM") ? 500 : (zys == "ERP") ? 300 : 0;
                }
                else if (amb == "QAS")
                {
                    mand = (zys == "CRM") ? 460 : (zys == "ERP") ? 260 : 0;
                }
                else if (amb == "DEV")
                {
                    mand = (zys == "CRM") ? 420 : (zys == "ERP") ? 120 : 0;
                }
                else
                {
                    mand = 0;
                }
            }
            return mand;
        }
        /// <summary>
        /// Método para Ingresar Ambiente y obtener el destination para una conexión a SAP sin crear la función ni sus parámetros.
        /// </summary>
        /// <param name="mandante">El número mandante de SAP: 120 (ERP-DEV), 420(CRM-DEV), 260(ERP-QA), 460(CRM-QA), 300(ERP-PRD), 500(CRM-PRD).</param>
        /// <returns>Retorna un objeto RfcDestination que es el destination.</returns>
        public RfcDestination GetDestRFC(string system, [Optional] int mandante)
        {
            RfcDestination destination = null;
            #region default
            mandante = checkDefault(system, mandante);
            if (mandante == 0)
            {
                return destination;
            }
            #endregion
            cred.IngresarAmbiente(mandante);
            destination = RfcDestinationManager.GetDestination(cred.parametros);

            return destination;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~SAP_Variants()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        void IDisposable.Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
