using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Net;
using DataBotV5.Data.Database;
using DataBotV5.Logical.Webex;
using DataBotV5.Data.Projects.PositionLetter;
using DataBotV5.Data.Stats;

namespace DataBotV5.Automation.RPA.HumanCapital
{
    /// <summary>
    /// Clase RPA Automation encargada de la carta de posición de human capital.
    /// </summary>
    class PositionLetter 
    {
        PositionLetterSQL position = new PositionLetterSQL();
        public void Main()
        {
            CreatePaper(position.Processes());
        }
        public void CreatePaper(DataTable process)
        {
            if (process.Rows.Count > 0)
            {
                for (int i = 0; i < process.Rows.Count; i++)
                {
                    string id = process.Rows[i][2].ToString();
                    string usuario = process.Rows[i][3].ToString();
                    string gestion = process.Rows[i][4].ToString();

                    Cartapos carta = JsonConvert.DeserializeObject<Cartapos>(gestion);
                    LetterWord(carta, id, usuario);
                    CRUD update = new CRUD();
                    //update.Update("Databot", "UPDATE condiciones_posicion SET ESTADO = 'COMPLETADO' WHERE ID_GESTION = '" + id + "'", "automation");
                }
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }
        public void LetterWord(Cartapos paper, string id, string user)
        {
            int tabla_id = 0;
            string fileName = @"C:\Users\databot02\Desktop\Databot\machote\carta.docx";
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application { Visible = false };
            Microsoft.Office.Interop.Word.Document aDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: true);
            aDoc.Activate();

            aDoc.Bookmarks["FECHA"].Select();
            wordApp.Selection.TypeText(DateTime.Now.ToString("dd/MM/yyyy"));
            aDoc.Bookmarks["COLABORADOR"].Select();
            wordApp.Selection.TypeText(paper.General.CollaboratorName);
            aDoc.Bookmarks["PAIS"].Select();
            wordApp.Selection.TypeText(paper.General.Country);
            aDoc.Bookmarks["UNIDAD"].Select();
            wordApp.Selection.TypeText(paper.General.Unit);
            aDoc.Bookmarks["FUNCION"].Select();
            wordApp.Selection.TypeText(paper.General.Funtion);
            aDoc.Bookmarks["POSICION"].Select();
            wordApp.Selection.TypeText(paper.General.Position);
            aDoc.Bookmarks["AREA"].Select();
            wordApp.Selection.TypeText(paper.General.Area);
            aDoc.Bookmarks["CECO"].Select();
            wordApp.Selection.TypeText(paper.General.CECO);
            aDoc.Bookmarks["DIRECTO"].Select();
            wordApp.Selection.TypeText(paper.Leadership.DirectBoss + " " + paper.Leadership.DirectBossPosition);
            if (paper.Leadership.ReportIndirect.Count > 0)
            {
                aDoc.Bookmarks["INDIRECTO"].Select();
                aDoc.Tables.Add(aDoc.Bookmarks["INDIRECTO"].Range, (paper.Leadership.ReportIndirect.Count + 1), 2);
                var tabla = aDoc.Tables[1];
                tabla.Cell(1, 1).Range.Text = "Nombre del jefe indirecto";
                tabla.Cell(1, 2).Range.Text = "Posición del jefe indirecto";
                for (int i = 0; i < paper.Leadership.ReportIndirect.Count; i++)
                {
                    int num = i + 2;
                    tabla.Cell(num, 1).Range.Text = paper.Leadership.ReportIndirect[i].Indirect_Boss;
                    tabla.Cell(num, 2).Range.Text = paper.Leadership.ReportIndirect[i].IndirectBossPosition;
                }
                tabla_id = tabla_id + 1;
            }
            aDoc.Bookmarks["LABOR"].Select();
            wordApp.Selection.TypeText(paper.General.DateLabor);
            if (paper.Leadership.PersonnelInCharge.Count > 0)
            {
                int index_tabla = 0;
                if (tabla_id != 0)
                {
                    index_tabla = 2;
                }
                else
                {
                    index_tabla = 1;
                }
                aDoc.Bookmarks["SUBALTERNOS"].Select();
                aDoc.Tables.Add(aDoc.Bookmarks["SUBALTERNOS"].Range, (paper.Leadership.PersonnelInCharge.Count + 1), 2);
                var tabla = aDoc.Tables[index_tabla];
                tabla.Cell(1, 1).Range.Text = "Nombre del colaborador";
                tabla.Cell(1, 2).Range.Text = "Posición del colaborador";
                for (int i = 0; i < paper.Leadership.PersonnelInCharge.Count; i++)
                {
                    int num = i + 2;
                    tabla.Cell(num, 1).Range.Text = paper.Leadership.PersonnelInCharge[i].Collaborator;
                    tabla.Cell(num, 2).Range.Text = paper.Leadership.PersonnelInCharge[i].Position;
                }
            }
            aDoc.Bookmarks["TERRITORIO"].Select();
            wordApp.Selection.TypeText(paper.Position.Territory);

            string lista_objetivos = "";
            for (int i = 0; i < paper.Position.Objetives.Count; i++)
            {
                lista_objetivos = lista_objetivos + "•	" + paper.Position.Objetives[i].Objetive + "\r\n";
            }

            aDoc.Bookmarks["OBJETIVOS"].Select();
            wordApp.Selection.TypeText(lista_objetivos);

            aDoc.Bookmarks["EVALUACION"].Select();
            wordApp.Selection.TypeText(paper.Position.Evaluation);

            string desc_asignacion = "";

            if (paper.Position.Period == "Permanente")
            {
                desc_asignacion = "La posición es permanente.";
            }
            else
            {
                desc_asignacion = "La posición es temporal, e inicia en la fecha " + paper.Position.DateStarAssignment + " finalizando en la fecha " + paper.Position.DateEndAssignment;
            }
            aDoc.Bookmarks["ASIGNACION"].Select();
            wordApp.Selection.TypeText(desc_asignacion);

            string compensacion = "";

            compensacion = compensacion + "•	Paquete mensual en moneda local: " + paper.Compensation.Salary + "\r\n";
            compensacion = compensacion + "•	Composición salarial: " + paper.Compensation.ComposicionFija + "% de salario fijo y " + paper.Compensation.CompositionVariable + " % de salario variable" + "\r\n";
            compensacion = compensacion + "•	El pago de su salario se hará en : " + paper.Compensation.Coin + "\r\n";
            compensacion = compensacion + "•	El salario fijo se le pagará mensualmente a partir de la fecha efectiva y el nuevo salario variable se le pagará según las políticas establecidas por la compañía, a partir de los resultados de hace dos meses." + "\r\n";
            compensacion = compensacion + "•	El pago del porcentaje del salario variable dependerá del grado de cumplimiento de los objetivos asignados en su nueva posición (según carta de objetivos o EPM)." + "\r\n";
            if (paper.Compensation.Protection != null)
            {
                if (paper.Compensation.Protection != "No aplica")
                {
                    compensacion = compensacion + "•	Como parte del proceso de toma de su nueva posición, se le protegerá el salario variable por un periodo de " + paper.Compensation.Protection + ", periodo en el cual usted recibirá un salario variable correspondiente al 100% durante ese periodo. Vencido el periodo de protección recibirá su pago basado en los resultados de cumplimiento de objetivos de dos (2) meses atrás." + "\r\n";
                }
            }

            aDoc.Bookmarks["COMPENSACION"].Select();
            wordApp.Selection.TypeText(compensacion);

            string generales = "";

            for (int i = 0; i < paper.Benefits.Generals.Count; i++)
            {
                generales = generales + "•	" + paper.Benefits.Generals[i] + "\r\n";
            }
            aDoc.Bookmarks["GENERALES"].Select();
            wordApp.Selection.TypeText(generales);
            string delapos = "";
            for (int i = 0; i < paper.Benefits.Position.Count; i++)
            {
                delapos = delapos + "•	" + paper.Benefits.Position[i] + "\r\n";
            }
            aDoc.Bookmarks["DELAPOS"].Select();
            wordApp.Selection.TypeText(delapos);

            string ruta_tmp = @"C:\Users\databot02\Desktop\Databot\machote\" + id + ".docx";

            aDoc.SaveAs2(ruta_tmp);
            aDoc.Close();
            wordApp.Quit();

            CreateFolderAndSave(id, id + ".docx", ruta_tmp, user);
            File.Delete(ruta_tmp);
        }

        public void CreateFolderAndSave(string folder, string file, string tempRoute, string user)
        {
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            string respuesta;
            string ruta = "ftp://10.7.60.72/carta_posiciones/" + folder;
            WebRequest request = WebRequest.Create(ruta);
            try
            {
                request.Method = WebRequestMethods.Ftp.MakeDirectory;
                request.Credentials = new NetworkCredential("gbmadmin", "P@ssword");
                using (var resp = (FtpWebResponse)request.GetResponse())
                {
                    respuesta = resp.StatusCode.ToString();
                }
            }
            catch (Exception)
            {
                //La ruta ya existe
            }
            using (var client = new WebClient())
            {
                client.Credentials = new NetworkCredential("gbmadmin", "P@ssword");
                client.UploadFile(ruta + "/" + file, WebRequestMethods.Ftp.UploadFile, tempRoute);
            }
            WebexTeams wb = new WebexTeams();
            string link = ruta + "/" + file;
            wb.SendLetterHCM(user, link);

        }
    }
    public class Cartapos
    {
        public Benefits Benefits { get; set; }
        public Compensation Compensation { get; set; }
        public General General { get; set; }
        public Leadership Leadership { get; set; }
        public PosicionCarta Position { get; set; }
        public string Id { get; set; }
        public string User { get; set; }
    }
    public class Benefits
    {
        public List<string> Generals { get; set; }
        public List<string> Position { get; set; }
    }
    public class Compensation
    {
        public string Salary { get; set; }
        public string ComposicionFija { get; set; }
        public string CompositionVariable { get; set; }
        public string Coin { get; set; }
        public string Protection { get; set; }
    }
    public class General
    {
        public string CollaboratorName { get; set; }
        public string Position { get; set; }
        public string Unit { get; set; }
        public string Area { get; set; }
        public string Country { get; set; }
        public string DateLabor { get; set; }
        public string Funtion { get; set; }
        public string CECO { get; set; }
    }
    public class Leadership
    {
        public string DirectBoss { get; set; }
        public string DirectBossPosition { get; set; }
        public List<Personal> PersonnelInCharge { get; set; }
        public List<IndirectBoss> ReportIndirect { get; set; }

    }
    public class Personal
    {
        public string Collaborator { get; set; }
        public string Position { get; set; }
    }
    public class IndirectBoss
    {
        public string Indirect_Boss { get; set; }
        public string IndirectBossPosition { get; set; }
    }
    public class PosicionCarta
    {
        public string Evaluation { get; set; }
        public string DateEndAssignment { get; set; }
        public string DateStarAssignment { get; set; }
        public string Period { get; set; }
        public string Territory { get; set; }
        public List<Objetives> Objetives { get; set; }
    }
    public class Objetives
    {
        public string Objetive { get; set; }
    }
}
