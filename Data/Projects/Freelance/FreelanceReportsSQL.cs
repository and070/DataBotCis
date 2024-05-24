using DataBotV5.Data.Database;
using DataBotV5.Logical.Mail;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataBotV5.Data.Projects.Freelance
{
    class FreelanceReportsSQL
    {
        //CRUD crud = new CRUD();
        //public DataTable Requests()
        //{
        //    DataTable data = new DataTable();
        //    string sql = "SELECT * FROM reportes_freelance WHERE ESTADO = 'EN PROCESO'";
        //    //data = crud.Select("Databot", sql, "automation");

        //    return data;
        //}
        //public void SendReport(string solicitante, string fileRoute, string id, string title)
        //{
        //    MailInteraction mail = new MailInteraction();
        //    mail.SendNotificationPortalFreelanceAnalytics(solicitante, title, fileRoute);
        //    string sql_update = "UPDATE reportes_freelance SET ESTADO = 'COMPLETADO' WHERE ID = '" + id + "'";
        //    CRUD cr = new CRUD();
        //    //cr.Update("Databot", sql_update, "automation");
        //    File.Delete(fileRoute);
        //}
        //public DataTable Extractor(string sql)
        //{
        //    DataTable data = new DataTable();
        //    //data = crud.Select("Databot", sql, "automation");
        //    return data;
        //}
    
    }
}
