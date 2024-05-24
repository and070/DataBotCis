using DataBotV5.Logical.Projects.Modals.Single;
using DataBotV5.Logical.Projects.ContactCenterOportunity;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace DataBotV5.Logical.Projects.GestionOpp
{
    /// <summary>
    /// Clase Logical encargada de la gestión de oportunidades.
    /// </summary>
    class GestionOpp
    {
        public string IDENTIFICADOR { get; set; }
        public string ID_GESTION { get; set; }
        public DG DATA_GENERAL { get; set; }
        public DORG DATA_ORG { get; set; }
        public List<DEQ> DATA_EQUIPO { get; set; }
        public List<LDRS> LDR { get; set; }
        public List<DBPM> BPM { get; set; }
        public string ESTADO { get; set; }
        public string EMPLEADO { get; set; }
        public string OPP { get; set; }
        public List<string> ARCHIVOS { get; set; }
        public string TS_CREACION { get; set; }
        public string TS_BPM { get; set; }
        public string TS_COMPLETADO { get; set; }
        public string TS_CRM { get; set; }
        public string TS_LDR { get; set; }
        public string SALES_OFFICE { get; set; }
        public string SALES_GROUP { get; set; }
        public string CORREO_EMP_RESPONSABLE { get; set; }
        public string NOMBRE_EMP_RESPONSABLE { get; set; }

        public GestionOpp(string identificador, string id_gestion, string dg, string dorg, string deq, string ldr, string bpm,
                          string estado, string empleado, string opp, string archivos, string ts1, string ts2, string ts3, string ts4, string ts5)
        {
            DG data_general = new DG();
            DORG data_organizacion = new DORG();
           
            
            DBPM requerimiento = new DBPM();
            List<DEQ> arrDEQ = new List<DEQ>();
            List<LDRS> arrLDR = new List<LDRS>();
            List<DBPM> arrBPM = new List<DBPM>();
            List<string> arrArchivos = new List<string>();
            IDENTIFICADOR = identificador;
            ID_GESTION = id_gestion;
            ESTADO = estado;
            //EMPLEADO = empleado;
            OPP = opp;
            ARCHIVOS = ARCHIVOS;
            TS_CREACION = ts1;
            TS_BPM = ts2;
            TS_COMPLETADO = ts3;
            TS_CRM = ts4;
            TS_LDR = ts5;

           
            JArray Jdg = JArray.Parse(dg);
            data_general.TIPO = Jdg[0]["TIPO"].Value<string>();
            data_general.DESCRIPCION = Jdg[0]["DESCRIPCION"].Value<string>();
            data_general.FECHA_INICIAL = Jdg[0]["FECHA_INICIAL"].Value<string>();
            data_general.FECHA_FINAL = Jdg[0]["FECHA_FINAL"].Value<string>();
            data_general.CICLO = Jdg[0]["CICLO"].Value<string>();
            data_general.ORIGEN = Jdg[0]["ORIGEN"].Value<string>();
            DATA_GENERAL = data_general;

            JArray Jdorg = JArray.Parse(dorg);
            data_organizacion.CLIENTE = Jdorg[0]["CLIENTE"].Value<string>();
            data_organizacion.CONTACTO = Jdorg[0]["CONTACTO"].Value<string>();
            data_organizacion.ORG_SERVICIOS = Jdorg[0]["ORG_SERVICIOS"].Value<string>();
            data_organizacion.ORG_VENTAS = Jdorg[0]["ORG_VENTAS"].Value<string>();

            DATA_ORG = data_organizacion;

            JArray Jdeq = JArray.Parse(deq);
            for (int i = 0; i < Jdeq.Count; i++)
            {
                DEQ data_equipo = new DEQ();
                data_equipo.ROL = Jdeq[i]["ROL"].Value<string>();
                data_equipo.EMP = Jdeq[i]["EMP"].Value<string>();
                arrDEQ.Add(data_equipo);
            }
            DATA_EQUIPO = arrDEQ;
            JArray Jldr = JArray.Parse(ldr);
            for (int i = 0; i < Jldr.Count; i++)
            {
                LDRS levanatamiento = new LDRS();
                List<ItemLDR> arrLisItem = new List<ItemLDR>();

                levanatamiento.TECNOLOGIA = Jldr[i]["TECNOLOGIA"].Value<string>();
                JObject items = JObject.Parse(Jldr[i].ToString());
                JArray ldrs = JArray.FromObject(items["LDR"]);
                for (int z = 0; z < ldrs.Count; z++)
                {
                    ItemLDR arrItem = new ItemLDR();
                    arrItem.ID = ldrs[z]["ID"].Value<int>();
                    arrItem.VALUE = ldrs[z]["VALUE"].Value<string>();
                    arrLisItem.Add(arrItem);
                }
                levanatamiento.LDR = arrLisItem;
                arrLDR.Add(levanatamiento);
            }
            LDR = arrLDR;
            JArray Jbpm = JArray.Parse(bpm);
            for (int i = 0; i < Jbpm.Count; i++)
            {
                DBPM data_bpm = new DBPM();
                data_bpm.PROVEEDOR = Jbpm[i]["PROVEEDOR"].Value<string>();
                data_bpm.PRODUCTO = Jbpm[i]["PRODUCTO"].Value<string>();
                data_bpm.REQUERIMIENTO = Jbpm[i]["REQUERIMIENTO"].Value<string>();
                data_bpm.CANTIDAD = Jbpm[i]["CANTIDAD"].Value<string>();
                data_bpm.INTEGRACION = Jbpm[i]["INTEGRACION"].Value<string>();
                data_bpm.COMENTARIOS = Jbpm[i]["COMENTARIOS"].Value<string>();
                arrBPM.Add(data_bpm);
            }
            BPM = arrBPM;

            CCEmployee data_empleado = new CCEmployee(empleado);
            EMPLEADO = data_empleado.IdEmpleado;
            CORREO_EMP_RESPONSABLE = data_empleado.Correo;
            NOMBRE_EMP_RESPONSABLE = data_empleado.Nombre;



            //< option value = "O 50000065" > CR </ option >
            //< option value = "O 50000077" > DO </ option >
            // < option value = "O 50000087" > SV </ option >
            //  < option value = "O 50000133" > HN </ option >
            //   < option value = "O 50000138" > MD </ option >
            //    < option value = "O 50000140" > NI </ option >
            //     < option value = "O 50000142" > PA </ option >
            //< option value = "O 50001429" > GT </ option >
            switch (data_organizacion.ORG_VENTAS)
            {
                case "O 50000065":
                    SALES_OFFICE = "OFF-CR";
                    SALES_GROUP = "DI_CR";
                    break;
                case "O 50000077":
                    SALES_OFFICE = "OFF-DO";
                    SALES_GROUP = "DI_DO";
                    break;
                case "O 50000087":
                    SALES_OFFICE = "OFF-SV";
                    SALES_GROUP = "DI_SV";
                    break;
                case "O 50000133":
                    SALES_OFFICE = "OFF-HN";
                    SALES_GROUP = "DI_HN";
                    break;
                case "O 50000138":
                    SALES_OFFICE = "OFF-MI";
                    SALES_GROUP = "DI_MI";
                    break;
                case "O 50000140":
                    SALES_OFFICE = "OFF-NI";
                    SALES_GROUP = "DI_NI";
                    break;
                case "O 50000142":
                    SALES_OFFICE = "OFF-PA";
                    SALES_GROUP = "DI_PA";
                    break;
                case "O 50001429":
                    SALES_OFFICE = "OFF-GT";
                    SALES_GROUP = "DI_GT";
                    break;
            }

            JArray Jarch = JArray.Parse(archivos);
            for (int i = 0; i < Jarch.Count; i++)
            {
                arrArchivos.Add(Jarch[i].Value<string>());
            }
            ARCHIVOS = arrArchivos;
        }
    }
}
