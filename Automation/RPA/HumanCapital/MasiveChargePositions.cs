using System;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using SAP.Middleware.Connector;
using System.Net;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Windows.Forms;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using System.Collections.Generic;

namespace DataBotV5.Automation.RPA.HumanCapital
{
    /// <summary>
    /// Clase RPA Automation encargada de la carga masiva de posiciones en Human Capital.
    /// </summary>
    class MasiveChargePositions 
    {
        #region variables globales
        Credentials cred = new Credentials();
        MailInteraction mail = new MailInteraction();
        Rooting root = new Rooting();
        ValidateData val = new ValidateData();
        ConsoleFormat console = new ConsoleFormat();
        ProcessInteraction proc = new ProcessInteraction();
        Log log = new Log();
        Stats estadisticas = new Stats();
        SapVariants sap = new SapVariants();
        public string response = "";
        public string response_failure = "";
        string nombre_de_la_posicion = "";
        string fecha_de_ingreso = "";
        string organizational_unit = "";
        string organizational_unit_id = "";
        string pais = "";
        string personal_subarea = "";
        string name_of_manager = "";
        string name_of_manager_id = "";
        string cost_center = "";
        string job = "";
        string job_id = "";
        string vacanty = "";
        string company_code = "";
        string personnel_area = "";
        string personnel_area_id = "";
        string personnel_subarea = "";
        string personnel_subarea_id = "";
        string admin_product = "";
        string admin_product_id = "";
        string direccion = "";
        string direccion_id = "";
        string epm = "";
        string fijo = "";
        string gerencia = "";
        string headcount = "";
        string headcount_id = "";
        string linea_de_negocio = "";
        string linea_de_negocio_id = "";
        string local_regional = "";
        string local_regional_pla = "";
        string pago_fijo = "";
        string pago_fijo_id = "";
        string pago_variable = "";
        string pago_variable_id = "";
        string productividad = "";
        string proteccion = "";
        string puesto_ccss = "";
        string puesto_ins = "";
        string recurso_de_inversion = "";
        string variable = "";
        string employee_group = "";
        string employee_group_id = "";
        string ee_subgroup = "";
        string ee_subgroup_id = "";
        string respuesta = "";
        string respuesta2 = "";
        string id_posicion = "";
        string Local_regional_salary = "";
        string devolver_value = "";
        string devolver_titulo = "";
        string Sociedad = "";
        string mandante = "ERP";
        string clipData = "";
        string respFinal = "";


        #endregion
        public void Main()
        {
            if (!sap.CheckLogin(mandante))
            {
                // leer correo y descargar archivo carga masiva posiciones
                if (mail.GetAttachmentEmail("Solicitudes Posicion Masiva", "Procesados", "Procesados Posicion Masiva"))
                {
                    sap.BlockUser(mandante, 1);
                    console.WriteLine("Procesando...");
                    ProcessMasivePositions(root.FilesDownloadPath + "\\" + root.ExcelFile);
                    response = "";
                    sap.BlockUser(mandante, 0);
                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }
                }
            }
        }
        public void ProcessMasivePositions(string route)
        {
            #region Variables Privadas
            int rows;
            string mensaje_devolucion = "";
            string validar_strc;
            bool validar_lineas = true;
            bool devolver = false;
            respuesta = "";
            respuesta2 = "";
            string validacion = "";
            #endregion

            #region abrir excel
            console.WriteLine("Abriendo Excel y Validando");
            Excel.Range xlRango;
            Excel.Range xlRangoDuplicate;

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(route);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];
            rows = xlWorkSheet.UsedRange.Rows.Count;
            #endregion

            validacion = xlWorkSheet.Cells[1, 26].text.ToString().Trim();

            if (validacion != "Local_regional_salary")
            {
                mensaje_devolucion = "Utilizar la plantilla oficial de masivo de posiciones";
                validar_lineas = false;
            }
            else
            {
                xlWorkSheet.Cells[1, 27].value = "ID de la posicion";
                xlWorkSheet.Cells[1, 28].value = "Resultado";
                xlWorkSheet.Range["A1"].Copy();
                Microsoft.Office.Interop.Excel.XlPasteType paste = Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats;
                Microsoft.Office.Interop.Excel.XlPasteSpecialOperation pasteop = Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationMultiply;
                xlWorkSheet.Range["AA1"].PasteSpecial(paste, pasteop, false, false);
                xlWorkSheet.Range["AB1"].PasteSpecial(paste, pasteop, false, false);
                for (int i = 2; i <= rows; i++)
                {
                    #region limpiar variables
                    id_posicion = "";
                    name_of_manager_id = "";
                    response = "";
                    response_failure = "";
                    nombre_de_la_posicion = "";
                    fecha_de_ingreso = "";
                    organizational_unit = "";
                    pais = "";
                    personal_subarea = "";
                    name_of_manager = "";
                    name_of_manager_id = "";
                    cost_center = "";
                    job = "";
                    vacanty = "";
                    company_code = "";
                    personnel_area = "";
                    personnel_subarea = "";
                    admin_product = "";
                    direccion = "";
                    epm = "";
                    fijo = "";
                    gerencia = "";
                    headcount = "";
                    linea_de_negocio = "";
                    local_regional = "";
                    local_regional_pla = "";
                    pago_fijo = "";
                    pago_variable = "";
                    productividad = "";
                    proteccion = "";
                    puesto_ccss = "";
                    puesto_ins = "";
                    recurso_de_inversion = "";
                    variable = "";
                    employee_group = "";
                    ee_subgroup = "";
                    respuesta = "";
                    id_posicion = "";
                    Local_regional_salary = "";
                    devolver_value = "";
                    devolver_titulo = "";

                    #endregion

                    nombre_de_la_posicion = xlWorkSheet.Cells[i, 1].text.ToString().Trim().ToUpper();
                    fecha_de_ingreso = xlWorkSheet.Cells[i, 2].text.ToString().Trim();
                    organizational_unit = xlWorkSheet.Cells[i, 3].text.ToString().Trim().ToUpper();
                    name_of_manager = xlWorkSheet.Cells[i, 4].text.ToString().Trim().ToUpper();
                    cost_center = xlWorkSheet.Cells[i, 5].text.ToString().Trim().ToUpper();
                    job = xlWorkSheet.Cells[i, 6].text.ToString().Trim().ToUpper();
                    company_code = xlWorkSheet.Cells[i, 7].text.ToString().Trim().ToUpper();
                    personnel_area = xlWorkSheet.Cells[i, 8].text.ToString().Trim().ToUpper();
                    personnel_subarea = xlWorkSheet.Cells[i, 9].text.ToString().Trim().ToUpper();
                    admin_product = xlWorkSheet.Cells[i, 10].text.ToString().Trim().ToUpper();
                    direccion = xlWorkSheet.Cells[i, 11].text.ToString().Trim().ToUpper();
                    epm = xlWorkSheet.Cells[i, 12].text.ToString().Trim().ToUpper();
                    fijo = xlWorkSheet.Cells[i, 13].text.ToString().Trim().ToUpper();
                    gerencia = xlWorkSheet.Cells[i, 14].text.ToString().Trim().ToUpper();
                    headcount = xlWorkSheet.Cells[i, 15].text.ToString().Trim().ToUpper();
                    linea_de_negocio = xlWorkSheet.Cells[i, 16].text.ToString().Trim().ToUpper();
                    local_regional_pla = xlWorkSheet.Cells[i, 17].text.ToString().Trim().ToUpper();
                    pago_fijo = xlWorkSheet.Cells[i, 18].text.ToString().Trim().ToUpper();
                    pago_variable = xlWorkSheet.Cells[i, 19].text.ToString().Trim().ToUpper();
                    productividad = xlWorkSheet.Cells[i, 20].text.ToString().Trim().ToUpper();
                    proteccion = xlWorkSheet.Cells[i, 21].text.ToString().Trim().ToUpper();
                    recurso_de_inversion = xlWorkSheet.Cells[i, 22].text.ToString().Trim().ToUpper();
                    variable = xlWorkSheet.Cells[i, 23].text.ToString().Trim().ToUpper();
                    employee_group = xlWorkSheet.Cells[i, 24].text.ToString().Trim().ToUpper();
                    ee_subgroup = xlWorkSheet.Cells[i, 25].text.ToString().Trim().ToUpper();
                    Local_regional_salary = xlWorkSheet.Cells[i, 26].text.ToString().Trim().ToUpper();


                    if (nombre_de_la_posicion == "")
                    {
                        continue;
                    }
                    else // si hay data
                    {
                        #region extraer los keys y validacion


                        if (company_code.Substring(0, 2) == "GB" || company_code.Substring(0, 2) == "LC")
                        {
                            pais = company_code.Substring(2, 2);
                            Sociedad = company_code.Substring(2, 2);
                        }
                        else
                        {
                            pais = company_code.Substring(0, 2);
                            Sociedad = company_code.Substring(0, 2);
                        }
                        if (Sociedad == "MD")
                        {
                            Sociedad = "MI";
                        }
                        cost_center = pais + cost_center.Substring(2, cost_center.Length - 2);

                        //necesario para que funcione mediante RFC de SAP (en SAP Gui no hace falta el CO01)
                        cost_center = cost_center + " CO01";

                        var valores = GetResponse("https://smartsimple.gbm.net:43888/bulk-load/find-all-information");
                        JObject obj = JObject.Parse(valores);

                        //key de unidad organizativa
                        var unidades = obj["payload"]["OrganizationalUnit"].ToString();
                        Pos[] unidad_key = JsonConvert.DeserializeObject<Pos[]>(unidades);
                        try
                        {
                            Pos UG_keys = unidad_key.FirstOrDefault(z => z.organizationalUnit == organizational_unit);
                            organizational_unit_id = UG_keys.keyOrganizationalUnit;
                        }
                        catch (Exception)
                        {
                            devolver_value = organizational_unit;
                            devolver_titulo = "Unidad Organizativa";
                        }


                        //key de personeel area
                        var PersonalArea = obj["payload"]["PersonalArea"].ToString();
                        Pos[] PersonalArea_key = JsonConvert.DeserializeObject<Pos[]>(PersonalArea);
                        try
                        {
                            Pos PA_keys = PersonalArea_key.FirstOrDefault(z => z.personalArea == personnel_area);
                            personnel_area_id = Sociedad + PA_keys.keyPersonalArea;
                        }
                        catch (Exception)
                        {
                            devolver_value = personnel_area;
                            devolver_titulo = "Area de Personal";
                        }

                        //key de direccion
                        var Direction = obj["payload"]["Direction"].ToString();
                        try
                        {
                            Pos Direction_keys = JsonConvert.DeserializeObject<Pos[]>(Direction).FirstOrDefault(z => z.direction == direccion);
                            direccion_id = Direction_keys.keyDirection;
                        }
                        catch (Exception)
                        {
                            devolver_value = direccion;
                            devolver_titulo = "Direccion";
                        }


                        //key de BussinessLine
                        var BussinessLine = obj["payload"]["BussinessLine"].ToString();
                        try
                        {
                            Pos BussinessLine_keys = JsonConvert.DeserializeObject<Pos[]>(BussinessLine).FirstOrDefault(z => z.bussinessLine == linea_de_negocio);
                            linea_de_negocio_id = BussinessLine_keys.keyBussinessLine;
                        }
                        catch (Exception)
                        {
                            devolver_value = linea_de_negocio;
                            devolver_titulo = "Linea de Negocio";
                        }


                        //key de Access
                        var Access = obj["payload"]["Access"].ToString();
                        try
                        {
                            Pos Access_keys = JsonConvert.DeserializeObject<Pos[]>(Access).FirstOrDefault(z => z.access == admin_product);
                            admin_product_id = Access_keys.keyAccess;
                        }
                        catch (Exception)
                        {
                            devolver_value = admin_product;
                            devolver_titulo = "Administrativo / Productivo";
                        }

                        //key de BudgetedResource
                        var BudgetedResource = obj["payload"]["BudgetedResource"].ToString();
                        try
                        {
                            Pos BudgetedResource_keys = JsonConvert.DeserializeObject<Pos[]>(BudgetedResource).FirstOrDefault(z => z.budgetedResource == headcount);
                            headcount_id = BudgetedResource_keys.keyBudgetedResource;
                        }
                        catch (Exception)
                        {
                            devolver_value = headcount;
                            devolver_titulo = "HeadCount";
                        }

                        //key de employeeSubGroup
                        var employeeSubGroup = obj["payload"]["employeeSubGroup"].ToString();
                        try
                        {
                            Pos employeeSubGroup_keys = JsonConvert.DeserializeObject<Pos[]>(employeeSubGroup).FirstOrDefault(z => z.employeeSubGroup == ee_subgroup);
                            ee_subgroup_id = employeeSubGroup_keys.keyEmployeeSubGroup;
                        }
                        catch (Exception)
                        {
                            devolver_value = ee_subgroup;
                            devolver_titulo = "Employee SubGroup";
                        }

                        //key de PositionType
                        var PositionType = obj["payload"]["PositionType"].ToString();
                        try
                        {
                            Pos PositionType_keys = JsonConvert.DeserializeObject<Pos[]>(PositionType).FirstOrDefault(z => z.positionType == employee_group);
                            employee_group_id = PositionType_keys.keyPositionType;
                        }
                        catch (Exception)
                        {
                            devolver_value = employee_group;
                            devolver_titulo = "Employee Group";
                        }

                        //key de PersonalBranch
                        var PersonalBranch = obj["payload"]["PersonalBranch"].ToString();
                        try
                        {
                            Pos PersonalBranch_keys = JsonConvert.DeserializeObject<Pos[]>(PersonalBranch).FirstOrDefault(z => z.personalBranch == personnel_subarea);
                            personnel_subarea_id = PersonalBranch_keys.keyPersonalBranch;
                        }
                        catch (Exception)
                        {
                            devolver_value = personnel_subarea;
                            devolver_titulo = "Sub Area de Personal";
                        }

                        //key de FixedPercent
                        var FixedPercent = obj["payload"]["FixedPercent"].ToString();
                        try
                        {
                            Pos FixedPercent_key_fijo = JsonConvert.DeserializeObject<Pos[]>(FixedPercent).FirstOrDefault(z => z.fixedPercent == pago_fijo);
                            pago_fijo_id = FixedPercent_key_fijo.keyFixedPercent;
                        }
                        catch (Exception)
                        {
                            devolver_value = pago_fijo;
                            devolver_titulo = "Pago Fijo";
                        }

                        var VariablePercent = obj["payload"]["VariablePercent"].ToString();
                        try
                        {
                            Pos VariablePercent_key_var = JsonConvert.DeserializeObject<Pos[]>(VariablePercent).FirstOrDefault(z => z.variablePercent == pago_variable);
                            pago_variable_id = VariablePercent_key_var.keyVariablePercent;
                        }
                        catch (Exception)
                        {
                            devolver_value = pago_variable;
                            devolver_titulo = "Pago Variable";
                        }


                        //posiciones ins y ccss
                        if (company_code == "GBCR")
                        {
                            var Positions = obj["payload"]["Positions"].ToString();
                            try
                            {
                                Pos Positions_key_var = JsonConvert.DeserializeObject<Pos[]>(Positions).FirstOrDefault(z => z.position == nombre_de_la_posicion);
                                puesto_ins = Positions_key_var.keyIns;
                                puesto_ccss = Positions_key_var.keyCcss;
                            }
                            catch (Exception)
                            {
                                devolver_value = nombre_de_la_posicion;
                                devolver_titulo = "Puesto del INS o de la CCSS";
                            }
                        }


                        if (local_regional_pla == "LOCAL")
                        {
                            local_regional_pla = "001";
                        }
                        else if (local_regional_pla == "REGIONAL")
                        {
                            local_regional_pla = "002";
                        }
                        else
                        {
                            devolver_value = local_regional_pla;
                            devolver_titulo = "Local Regional PLA";
                        }

                        if (Local_regional_salary == "LOCAL")
                        {
                            Local_regional_salary = "001";
                        }
                        else if (Local_regional_salary == "REGIONAL")
                        {
                            Local_regional_salary = "002";
                        }
                        else
                        {
                            devolver_value = Local_regional_salary;
                            devolver_titulo = "Local Regional Salary";
                        }

                        epm = SiOrNo(epm);
                        if (epm == "")
                        { devolver_titulo = "EPM"; }

                        productividad = SiOrNo(productividad);
                        if (productividad == "")
                        { devolver_titulo = "Productividad"; }

                        proteccion = SiOrNo(proteccion);
                        if (proteccion == "")
                        { devolver_titulo = "proteccion"; }

                        recurso_de_inversion = SiOrNo(recurso_de_inversion);
                        if (recurso_de_inversion == "")
                        { devolver_titulo = "recurso_de_inversion"; }

                        gerencia = SiOrNo(gerencia);
                        if (gerencia == "")
                        { devolver_titulo = "Gerencia"; }

                        if (fijo == "COSTO")
                        {
                            fijo = "002";
                        }
                        else if (fijo == "GASTO")
                        {
                            fijo = "001";
                        }
                        else
                        {
                            devolver_titulo = "Fijo";
                        }

                        if (variable == "COSTO")
                        {
                            variable = "002";
                        }
                        else if (variable == "GASTO")
                        {
                            variable = "001";
                        }
                        else
                        {
                            devolver_titulo = "Variable";
                        }


                        try
                        {
                            Dictionary<string, string> parameters = new Dictionary<string, string>();
                            parameters["USERNAME"] = name_of_manager.ToUpper();

                            IRfcFunction func = sap.ExecuteRFC(mandante, "ZHR_GET_INFO_POSITION", parameters);


                            name_of_manager_id = func.GetValue("POSITION_USER").ToString();
                            if (name_of_manager_id == "")
                            {
                                xlWorkSheet.Cells[i, 28].value = "No se encontro la posicion del Manager: " + name_of_manager;
                                xlWorkSheet.Range["A" + i].Copy();
                                xlWorkSheet.Range["AB" + i].PasteSpecial(paste, pasteop, false, false);
                                validar_lineas = false;
                                continue;
                            }
                        }
                        catch (Exception ex)
                        {
                            response_failure = val.LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, i);
                            console.WriteLine(" Finishing process " + response_failure);
                            respuesta = nombre_de_la_posicion + ": " + ex.ToString();
                            response_failure = ex.ToString();
                            validar_lineas = false;
                            xlWorkSheet.Cells[i, 28].value = "No se encontro la posicion del Manager: " + name_of_manager;
                            xlWorkSheet.Range["A" + i].Copy();
                            xlWorkSheet.Range["AB" + i].PasteSpecial(paste, pasteop, false, false);
                            continue;
                        }


                        //valida que todos los ID esten:
                        if (devolver_value != "" || devolver_titulo != "")
                        {
                            xlWorkSheet.Cells[i, 28].value = "No se encontro el key de " + devolver_titulo + ": " + devolver_value;
                            xlWorkSheet.Range["A" + i].Copy();
                            xlWorkSheet.Range["AB" + i].PasteSpecial(paste, pasteop, false, false);
                            validar_lineas = false;
                            continue;
                        }


                        #endregion

                        #region SAP
                        console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);

                        try
                        {
                            Dictionary<string, string> parameters = new Dictionary<string, string>();
                            parameters["NOMBRE_DE_LA_POSICION"]= nombre_de_la_posicion;
                            parameters["FECHA_DE_INGRESO"] = fecha_de_ingreso;
                            parameters["ORGANIZATIONAL_UNIT"] = organizational_unit_id;
                            parameters["NAME_OF_MANAGER"] = name_of_manager_id;
                            parameters["COST_CENTER"] = cost_center;
                            parameters["JOB"] = job;
                            parameters["COMPANY_CODE"] = company_code;
                            parameters["PERSONNEL_AREA"] = personnel_area_id;
                            parameters["PERSONNEL_SUBAREA"] = personnel_subarea_id;
                            parameters["ADMIN_PRODUCT"] = admin_product_id;
                            parameters["DIRECCION"] = direccion_id;
                            parameters["EPM"] = epm;
                            parameters["FIJO"]= fijo;
                            parameters["GERENCIA"]= gerencia;
                            parameters["HEADCOUNT"]= headcount_id;
                            parameters["LINEA_DE_NEGOCIO"]= linea_de_negocio_id;
                            parameters["LOCAL_REGIONAL_PLA"]= local_regional_pla;
                            parameters["LOCAL_REGIONAL_SALARIAL"]= Local_regional_salary;
                            parameters["PAGO_FIJO"]= pago_fijo_id;
                            parameters["PAGO_VARIABLE"]= pago_variable_id;
                            parameters["PRODUCTIVIDAD"]= productividad;
                            parameters["PROTECCION"]= proteccion;
                            parameters["RECURSO_DE_INVERSION"]= recurso_de_inversion;
                            parameters["VARIABLE"]= variable;
                            parameters["EMPLOYEE_GROUP"]= employee_group_id;
                            parameters["EE_SUBGROUP"]= ee_subgroup_id;
                            parameters["PUESTO_CCSS"]= puesto_ccss;
                            parameters["PUESTO_INS"]= puesto_ins;
                            parameters["VACANTY"]= "0";

                            IRfcFunction func = sap.ExecuteRFC(mandante, "ZHR_POSITION_CREATE", parameters);

                            if (func.GetValue("RESPUESTA").ToString() == "NA JOB")
                            {
                                if (Local_regional_salary == "001")
                                { local_regional = "Local"; }
                                else
                                { local_regional = "Regional"; }
                                respuesta = "No se encontro el Job seleccionado: " + job + ". En el pais: " + company_code + ". A nivel: " + local_regional;
                            }
                            else if (func.GetValue("RESPUESTA").ToString() == "NA Unidad Z")
                            { respuesta = "No se encontro la Unidad Z homologa de la unidad: " + organizational_unit; }
                            else if (func.GetValue("RESPUESTA").ToString().Contains("Error:"))
                            { respuesta = func.GetValue("RESPUESTA").ToString(); }
                            else if (func.GetValue("RESPUESTA").ToString() == "OK")
                            {
                                respuesta = "Posicion creada con exito";
                                id_posicion = func.GetValue("ID_POSICION").ToString();
                            }
                            else
                            {
                                respuesta = "Error insesperado, contacte a Datos Maestros";
                                id_posicion = func.GetValue("ID_POSICION").ToString();
                            }
                            if (id_posicion != "")
                            {
                                clipData = clipData + id_posicion + "\r\n";
                            }
                            Console.WriteLine(DateTime.Now + " > > >  " + id_posicion + " : " + respuesta);
                            //log de base de datos
                            log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Delivery", id_posicion + " : " + respuesta, root.Subject);
                            respFinal = respFinal + "\\n" + id_posicion + " : " + respuesta;

                            xlWorkSheet.Cells[i, 27].value = id_posicion;
                            xlWorkSheet.Cells[i, 28].value = respuesta;
                            xlWorkSheet.Range["A" + i].Copy();
                            xlWorkSheet.Range["AA" + i].PasteSpecial(paste, pasteop, false, false);
                            xlWorkSheet.Range["AB" + i].PasteSpecial(paste, pasteop, false, false);
                        }
                        catch (Exception ex)
                        {
                            response_failure = val.LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, i);
                            console.WriteLine(" Finishing process " + response_failure);
                            respuesta = id_posicion + ": " + ex.ToString();
                            response_failure = ex.ToString();
                            validar_lineas = false;
                        }

                        #endregion

                    }
                } //for

                xlApp.DisplayAlerts = false;
                xlWorkBook.SaveAs(route);
                xlWorkBook.Close();

            } //else de validation
            xlApp.Workbooks.Close();
            xlApp.Quit();
            proc.KillProcess("EXCEL",true);

            console.WriteLine("Crear en tabla T");
            #region Crear en tabla T528 B/T
            try
            {
                if (clipData != "")
                {
                    Clipboard.SetText(clipData.ToString(), TextDataFormat.Text);
                    sap.LogSAP(mandante.ToString());
                    try
                    {
                        ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nse38";
                        SapVariants.frame.SendVKey(0);
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtRS38M-PROGRAMM")).Text = "RHINTE10";
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();

                        ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/chkMOVEOBJ")).Selected = true;
                        ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/chkTEST")).Selected = false;

                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtPCHPLVAR")).Text = "01";
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtPCHOTYPE")).Text = "S";
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txtPCHSEARK")).Text = "";
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtPCHOBJID-LOW")).Text = "";
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtPCHOSTAT")).Text = "1";
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtPCHISTAT")).Text = "1";

                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/usr/btn%_PCHOBJID_%_APP_%-VALU_PUSH")).Press();
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[24]")).Press();
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[8]")).Press();
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                    }
                    catch (Exception)
                    {
                        response_failure = "Error al crear ID en T528B";
                        validar_lineas = false;
                    }

                    sap.KillSAP();
                    Clipboard.Clear();
                }
            }
            catch (Exception ex)
            {
                response_failure = "Error al crear ID en T528B";
                validar_lineas = false;
                proc.KillProcess("saplogon",false);
                Clipboard.Clear();
            }

            #endregion

            console.WriteLine("Respondiendo solicitud");

            if (validar_lineas == false)
            {
                string[] adjunto = { root.FilesDownloadPath + "\\" + root.ExcelFile };
                //enviar email de repuesta de error
                string[] cc = { "appmanagement@gbm.net" };
                mail.SendHTMLMail("Los resultados estan en el excel" + "<br>" + response_failure, new string[] { root.BDUserCreatedBy }, root.Subject, cc, adjunto);
            }
            else
            {
                string[] adjunto = { root.FilesDownloadPath + "\\" + root.ExcelFile };
                //enviar email de repuesta de exito
                mail.SendHTMLMail("Los resultados estan en el excel", new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC, adjunto);
            }

            root.requestDetails = respFinal;

        }

        public static string GetResponse(string endPoint)
        {
            System.Net.HttpWebRequest request = CreateWebRequest(endPoint);

            using (var response = (HttpWebResponse)request.GetResponse())
            {
                var responseValue = string.Empty;

                if (response.StatusCode != HttpStatusCode.OK)
                {
                    string message = String.Format("POST failed. Received HTTP {0}", response.StatusCode);
                    throw new ApplicationException(message);
                }

                using (var responseStream = response.GetResponseStream())
                {
                    using (var reader = new StreamReader(responseStream))
                    {
                        responseValue = reader.ReadToEnd();
                    }
                }

                return responseValue;
            }
        }
        private static HttpWebRequest CreateWebRequest(string endPoint)
        {
            var request = (HttpWebRequest)WebRequest.Create(endPoint);

            request.Method = "GET";
            request.ContentLength = 0;
            request.ContentType = "text/json";
            request.Timeout = 90000;

            return request;
        }
        public string SiOrNo(string Value)
        {
            string key = "";
            if (Value == "SI")
            {
                key = "001";
            }
            else if (Value == "NO")
            {
                key = "002";
            }
            else
            {
                key = "";
            }

            return key;
        }
    }
    public class Pos
    {
        public string organizationalUnit { get; set; }
        public string keyOrganizationalUnit { get; set; }
        public string personalArea { get; set; }
        public string keyPersonalArea { get; set; }
        public string direction { get; set; }
        public string keyDirection { get; set; }
        public string bussinessLine { get; set; }
        public string keyBussinessLine { get; set; }
        public string access { get; set; }
        public string keyAccess { get; set; }
        public string budgetedResource { get; set; }
        public string keyBudgetedResource { get; set; }
        public string employeeSubGroup { get; set; }
        public string keyEmployeeSubGroup { get; set; }
        public string positionType { get; set; }
        public string keyPositionType { get; set; }
        public string requestType { get; set; }
        public string keyRequestType { get; set; }
        public string personalBranch { get; set; }
        public string keyPersonalBranch { get; set; }
        public string fixedPercent { get; set; }
        public string keyFixedPercent { get; set; }
        public string variablePercent { get; set; }
        public string keyVariablePercent { get; set; }
        public string position { get; set; }
        public string keyIns { get; set; }
        public string keyCcss { get; set; }
    }
 
}
