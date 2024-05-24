using DataBotV5.App.Global;
using DataBotV5.Data.Database;
using DataBotV5.Data.Projects.CrBids;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.Projects.CrBids;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;

namespace DataBotV5.Automation.MASS.CrBids
{
    /// <summary>
    /// Clase MASS Automation "Robot 4" encargada de mover concursos SICOP que no se participa porque la fecha de apertura ya pasó. 
    /// </summary>
    class PendingBidSS
    {
        #region variables_globales
        string enviroment = "QAS";
        CRUD crud = new CRUD();
        CrBidsLogical cr_licitaciones = new CrBidsLogical();
        ConsoleFormat console = new ConsoleFormat();
        Stats estadisticas = new Stats();
        BidsGbCrSql sqlCrBids = new BidsGbCrSql();
        MailInteraction mail = new MailInteraction();

        Rooting root = new Rooting();
        Log log = new Log();

        string respFinal = "";



        internal CrBidsLogical Cr_licitaciones { get => cr_licitaciones; set => cr_licitaciones = value; }
        internal Stats Estadisticas { get => estadisticas; set => estadisticas = value; }
        internal BidsGbCrSql Lcsql { get => sqlCrBids; set => sqlCrBids = value; }
        #endregion

        public void Main()
        {
            NoParticipa();
        }

        /// <summary>
        /// Este método se encarga de mover a backup todas las licitaciones que no han sido respondidas por el Account Manager
        /// donde indique si va a participar en la licitación o no, entonces el robot verifica si la licitacion no ha sido respondida y si ya 
        /// se cumplió la fecha de apertura y si es así mueve todo a backup . Todo esto partiendo de la tabla de purchaseOrders de Costa_Rica_Bids de Smart And Simple.
        /// </summary>
        public void NoParticipa()
        {
            console.WriteLine("Extrayendo las PurchaseOrders vencidas de DB");
            #region Extracción de información de BD.
            //Extrae sólo las PO vencidas y que el campo Participation no sea 'SI'.
            string sqlPurchaseOrders = "SELECT * FROM `purchaseOrder` WHERE offerOpening < CURRENT_TIMESTAMP AND id in (SELECT bidNumber FROM `purchaseOrderAdditionalData` WHERE participation!='SI' AND participation!='NO')";
            DataTable purchaseOrders = crud.Select( sqlPurchaseOrders, "costa_rica_bids_db");
            #endregion

            #region Variables Locales
            int cantPurchaseOrders = purchaseOrders.Rows.Count;
            Dictionary<string, string> BidsWithErrors = new Dictionary<string, string>();
            #endregion

            //Verifica si hay purchaseOrders a validar en Database.
            if (cantPurchaseOrders > 0)
            {
                console.WriteLine("Inicio de proceso de mover PurchaseOrders vencidas al backup.");

                //Recorrer c/u de las purchaseOrders vencidas (establecidas por el select).
                for (int i = 0; i < cantPurchaseOrders; i++)
                {
                    console.WriteLine("Bid #" + (i + 1) + " de " + cantPurchaseOrders + ", " +
                        "Bid Number: " + purchaseOrders.Rows[i]["bidNumber"].ToString() +
                        "  Institution: " + purchaseOrders.Rows[i]["institution"].ToString());

                    string idPurchaseOrder = purchaseOrders.Rows[i]["Id"].ToString();
                    bool resultMovement = sqlCrBids.MoveBidToBackup(idPurchaseOrder);

                    //Registra si existió un problema con el movimiento a backup de la licitación y lo agrega
                    //a un diccionario para enviar posteriormentecorreo con los errores.
                    if (!resultMovement)
                        BidsWithErrors.Add(idPurchaseOrder.ToString(), purchaseOrders.Rows[i]["institution"].ToString());
                    log.LogDeCambios("Modificación", root.BDProcess, "NA", "No participa por superar la fecha de apertura de ofertas", purchaseOrders.Rows[i]["bidNumber"].ToString(), "");
                    respFinal = respFinal + "\\n" + "Establece que no participa por superar la fecha de apertura de ofertas: " + purchaseOrders.Rows[i]["bidNumber"].ToString();

                }

                NotifyErrorsPendingBid(BidsWithErrors);

                root.requestDetails = respFinal;
                root.BDUserCreatedBy = "MGARCIA";


                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
                console.WriteLine("Proceso finalizado, todas las PurchaseOrders vencidas movidas al backup correctamente.");
            }
            else
            {
                console.WriteLine("Todo al día, no hay PurchaseOrders vencidas ó que no participen para mover a backup.");
            }

        }

        /// <summary>
        /// Método encargado de notificar mediante correo electrónico a los responsables del robot los registros de 
        /// licitaciones que no pudieron respaldarse al backup de la licitación, estos registros están almacenados en 
        /// un diccionario <id> <institution>, con esto posteriormente se notifican las licitaciones con problemas.
        /// </summary>
        /// <param name="BidsWithErrors"></param>
        void NotifyErrorsPendingBid(Dictionary<string, string> BidsWithErrors)
        {
            if (BidsWithErrors.Count > 0)
            {
                string msgEmail = 
                     "Robot: <strong>PendingBid</strong> <br>" +
                     "Proyecto: <strong>CrBids</strong> <br>" +
                     "Área: <strong>MASS</strong> <br><br>" +
                     "Las siguientes licitaciones debido a un error no se pudieron mover a backup:<br><br>";

                console.WriteLine("Las siguientes licitaciones no se han respaldado correctamente, por favor revisarlas: ");
                foreach (KeyValuePair<string, string> bid in BidsWithErrors)
                {
                    console.WriteLine($"Id: {bid.Key} - Institution: {bid.Value}") ;
                    msgEmail += $"Id: {bid.Key} - Institution: {bid.Value}<br> ";
                }
                msgEmail += $"<br><br>Por favor, proceder a revisarlas y verificar por qué no se realizaron los backups correctamente.<br><br>";

                mail.SendHTMLMail(msgEmail, new string[] {"appmanagement@gbm.net"}, "He registrado un error - Robot PendingBid" ,  new string[] { "epiedra@gbm.net", "dmeza@gbm.net" });
            }
        }

    }
}
