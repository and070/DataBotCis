using System;
using System.Data;
using MySql.Data.MySqlClient;
using DataBotV5.Data.Database;
using DataBotV5.Data.SAP;
using DataBotV5.App.Global;

namespace DataBotV5.Logical.Projects.UpdateClients
{
    /// <summary>
    /// Clase logical diseñada para actualizar vendedores, sales_rep, de fabrica de ofertas y licitaciones_cr.
    /// </summary>
    class UpdateClients
    {

        ConsoleFormat console = new ConsoleFormat();
        DataTable mytable = new DataTable();
        CRUD crud = new CRUD();
        Settings set = new Settings();
        SapVariants sap = new SapVariants();

        public string UpdateClient(string client, string salesRep)
        {

            try
            {
                if (salesRep.Substring(0, 2).ToUpper() != "AA")
                {
                    salesRep = "AA" + salesRep.PadLeft(8, '0');
                }

                client = client.TrimStart('0');

                string SalesRepUser = "";
                string manager = "";

                if (salesRep != "")
                {
                    #region Extraer usuario del vendedor

                    DataTable empleados = new DataTable();
                    try
                    {
                        string sql_select = "SELECT * FROM `digital_sign`";
                        empleados = crud.Select(sql_select, "MIS");

                        if (empleados.Rows.Count > 0)
                        {
                            System.Data.DataRow[] emp_info = empleados.Select("UserID ='" + salesRep.ToString() + "'"); //like '%" + institu + "%'"
                            SalesRepUser = emp_info[0]["user"].ToString().ToUpper();
                        }

                    }
                    catch (Exception ex)
                    {
                        console.WriteLine(ex.ToString());

                    }
                    if (string.IsNullOrEmpty(SalesRepUser))
                    {
                        using (SapVariants sap = new SapVariants())
                        {
                            SalesRepUser = sap.AmEmail(salesRep).Split('@')[0];
                        }
                    }

                    #endregion
                    #region GET Manager
                    DataTable row = new DataTable();
                    try
                    {
                        string sql = $"SELECT * FROM `clients` where accountManagerUser = '{SalesRepUser}'";
                        row = crud.Select(sql, "databot_db");

                        if (row.Rows.Count > 0)
                        {
                            manager = row.Rows[0]["manager"].ToString();
                        }

                    }
                    catch (Exception ex)
                    {
                        console.WriteLine(ex.ToString());

                    }
                    #endregion
                    salesRep = salesRep.ToString().TrimStart('A');
                }

                string query = $@"UPDATE 
                                        `clients` SET 
                                        `accountManagerId`= '{salesRep}' , 
                                        `accountManagerUser`= '{SalesRepUser}', 
                                        employeeResponsible`= (SELECT MIS.digital_sign.id from MIS.digital_sign WHERE MIS.digital_sign.UserID = '{SalesRepUser}'),
                                        `manager` = '{manager}'
                                        WHERE idClient = '{client}'";

                crud.Update(query, "databot_db");
                return "OK";

            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                return "ERR";
            }
        }
        /// <summary>
        /// Actualiza los SalesRepresentative de CostaRicaBids
        /// </summary>
        /// <param name="client"></param>
        /// <param name="salesRep"></param>
        /// <returns></returns>
        public string UpdateEntytiesOldCrAndPaBids(string client, string salesRep)
        {
            //if (salesRep.Substring(0, 2).ToUpper() != "AA")
            //{
            //    salesRep = "AA" + salesRep.PadLeft(8, '0');
            //}

            //try
            //{
            //    MySqlConnection conn = new Database().Conn("licitaciones_cr");
            //    string query = "UPDATE `entidades` SET `sales_rep` = '" + salesRep + "' WHERE cliente = '" + client.PadLeft(10, '0') + "'";
            //    conn.Open();


            //    MySqlCommand execute = new MySqlCommand(query, conn);
            //    int result = execute.ExecuteNonQuery();
            //    conn.Close();

            //    if (result == 0)
            //    {
            //        return "ERR";
            //    }
            //    else
            //    {
            //        string query2 = "UPDATE `concursos` SET `AM` = '" + salesRep + "' WHERE cliente_institucion = '" + client.PadLeft(10, '0') + "'";
            //        conn.Open();
            //        MySqlCommand execute2 = new MySqlCommand(query2, conn);
            //        int result2 = execute2.ExecuteNonQuery();
            //        conn.Close();
            //        if (result2 == 0)
            //        {
            //            return "ERR";
            //        }
            //        else
            //        {
            //            return "OK";
            //        }
            //    }
            //}
            //catch (Exception)
            //{
            //    return "ERR";
            //}
            return "OK";
        }



        /// <summary>
        /// Actualiza los SalesRepresentative asociados a un Cliente en la base de datos de CostaRicaBids 
        /// y PanamáBids.
        /// </summary>
        /// <param name="client"></param>
        /// <param name="salesRep"></param>
        /// <returns></returns>
        public void UpdateEntyties(string client, string salesRep, string enviroment)
        {

            string sqlUser = $"SELECT user FROM `digital_sign` WHERE `UserID` = {salesRep}";
            DataTable userSalesRepDt = crud.Select( sqlUser, "MIS");
            string userSalesRep = "";

            try
            {
                userSalesRep = userSalesRepDt.Rows[0]["user"].ToString();
            }
            catch (Exception e)
            {
                //Se busca en SAP el UserName
                try
                {
                    string auxUser = "AA"+ salesRep.PadLeft(8, '0');
                    userSalesRep = sap.AmEmail(auxUser).Split('@')[0];
                }
                catch (Exception i) //No lo encontró en SAP.
                {
                    string msgError = $"No se pudo actualizar el SalesRepresentative en DB debido a que no existe en SAP.<br>" +
                    $"Client={client}<br>SalesRep:{salesRep}";
                    set.SendError("UpdateClients", "No existe el SalesRepresentative en SAP", msgError, i);
                    return;
                }

            };


            #region Actualizar SalesRepresentative en CrBids.
            string sqlInsti = $"SELECT * FROM `institutions` WHERE `customerId` = '{client}'";
            DataTable costumerCrBids = crud.Select( sqlInsti, "costa_rica_bids_db");

            //Si existe el cliente en CrBids.
            if (costumerCrBids.Rows.Count > 0)
            {                
                string sqlUpdateInsti = $"UPDATE `institutions` SET `salesRepresentative`='{userSalesRep}' WHERE customerId = '{client}'";
                crud.Update(sqlUpdateInsti, "costa_rica_bids_db");


                string sqlUpdateAM = $"UPDATE `purchaseorderadditionaldata` SET `accountManager`= '{userSalesRep}' WHERE customerInstitute = '{client}'";
                crud.Update(sqlUpdateAM, "costa_rica_bids_db");                           

            }

            #endregion

            #region Actualizar SalesRepresentative en PaBids.
            string sqlEnti = $"SELECT * FROM `entities` WHERE `customerId` = '{client}';";
            DataTable costumerPaBids = crud.Select( sqlEnti, "panama_bids_db");

            //Si existe el cliente en PaBids.
            if (costumerPaBids.Rows.Count > 0)
            {
                string sqlUpdateInsti = $"UPDATE `entities` SET `salesRep`='{userSalesRep}' WHERE customerId = '{client}'";
                crud.Update(sqlUpdateInsti, "panama_bids_db");
            }

            #endregion

            #region Actualizar Cliente en DatabotDB.
            string sqlUpdateClient = $@"UPDATE 
                                        `clients` SET 
                                        `accountManagerId`= '{salesRep}' , 
                                        `accountManagerUser`= '{userSalesRep}', 
                                        employeeResponsible`= (SELECT MIS.digital_sign.id from MIS.digital_sign WHERE MIS.digital_sign.UserID = '{userSalesRep}')
                                        WHERE idClient = '{client}'";
            crud.Update(sqlUpdateClient, "databot_db");

            #endregion
        }
    }
}
