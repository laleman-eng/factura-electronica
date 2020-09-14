using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Configuration;
using System.Threading;
using System.Data.SqlClient;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Net;
using System.Linq;
using System.Data;
using System.IO;
using System.Xml;
using System.Drawing;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using SAPbouiCOM;
using SAPbobsCOM;
using VisualD.GlobalVid;
using VisualD.SBOFunctions;
using VisualD.vkBaseForm;
using VisualD.MultiFunctions;
using VisualD.vkFormInterface;
using VisualD.SBOObjectMg1;
using VisualD.Main;
using VisualD.MainObjBase;
using VisualD.ADOSBOScriptExecute;
using Factura_Electronica_VK.Functions;
using Newtonsoft.Json;

namespace Factura_Electronica_VK.EnviarEstadoDTE
{
    public class TEnviarEstadoDTE : TvkBaseForm, IvkFormInterface
    {
        private List<string> Lista;
        private SAPbobsCOM.Recordset oRecordSet;
        private SAPbouiCOM.DBDataSource oDBDSHeader;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.ComboBox oComboBox;
        private SAPbouiCOM.Grid oGrid;
        private SAPbouiCOM.Grid oGrid2;
        private SAPbouiCOM.DataTable oDataTable;
        private SAPbouiCOM.DataTable oDTParams;
        private SAPbouiCOM.DataTable oDTSQL;
        private SAPbouiCOM.DataTable oDTRES;
        private SAPbouiCOM.DBDataSource oDBDSHC;
        private SAPbouiCOM.DBDataSource oDBDSHV;
        private SAPbouiCOM.UserDataSource oUD_AC;
        private SAPbouiCOM.UserDataSource oUD_RC;
        private SAPbouiCOM.UserDataSource oUD_LN;
        private TFunctions Funciones = new TFunctions();
        private CultureInfo _nf = new System.Globalization.CultureInfo("en-US");
        private String s;
        private String TempDocNumLink;
        private int TotDocu = 0;
        private int TotAcep = 0;
        private int TotRecl = 0;
        private int TotSele = 0;

        public new bool InitForm(string uid, string xmlPath, ref Application application, ref SAPbobsCOM.Company company, ref CSBOFunctions SBOFunctions, ref TGlobalVid _GlobalSettings)
        {
            Int32 CantRol;

            bool Result = base.InitForm(uid, xmlPath, ref application, ref company, ref SBOFunctions, ref _GlobalSettings);

            oRecordSet = (SAPbobsCOM.Recordset)(FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
            Funciones.SBO_f = FSBOf;
            try
            {
                Lista = new List<string>();
                FSBOf.LoadForm(xmlPath, "VID_EnviarEstadoDTE.srf", uid);
                //EnableCrystal := true;
                oForm = FSBOApp.Forms.Item(uid);
                oForm.Visible = false;
                oForm.Freeze(true);
                oForm.AutoManaged = false;
                oForm.SupportedModes = -1;       // afm_All
                oForm.EnableMenu("1282", false); //Crear
                oForm.EnableMenu("1281", false); //Actualizar

                // Ok Ad  Fnd Vw Rq Sec
                //Lista.Add('DocNum    , f,  f,  t,  f, n, 1');
                //Lista.Add('DocDate   , f,  t,  f,  f, r, 1');
                //Lista.Add('CardCode  , f,  t,  t,  f, r, 1');
                //FSBOf.SetAutoManaged(var oForm, Lista);

                oDBDSHC = oForm.DataSources.DBDataSources.Add("@VID_FEDTECPRA");

                oDataTable = oForm.DataSources.DataTables.Add("dt");
                oDTParams = oForm.DataSources.DataTables.Add("DT_P");
                oDTSQL = oForm.DataSources.DataTables.Add("DT_SQL");
                oGrid = (Grid)(oForm.Items.Item("grid").Specific);
                oGrid.DataTable = oDataTable;
                oGrid.SelectionMode = BoMatrixSelect.ms_Single;

                oUD_AC = oForm.DataSources.UserDataSources.Item("UD_AC");
                oUD_AC.ValueEx = "Y";
                oUD_RC = oForm.DataSources.UserDataSources.Item("UD_RC");
                oUD_RC.ValueEx = "Y";
                oUD_LN = oForm.DataSources.UserDataSources.Item("UD_LN");
                oUD_LN.ValueEx = "Y";


                oDTRES = oForm.DataSources.DataTables.Add("DT_RES");
                oDTRES.Columns.Add("Descripcion", BoFieldsType.ft_AlphaNumeric, 254);
                oDTRES.Columns.Add("Documento", BoFieldsType.ft_AlphaNumeric, 50);
                oDTRES.Columns.Add("Fecha / Hora", BoFieldsType.ft_AlphaNumeric, 50);
                oGrid2 = (Grid)(oForm.Items.Item("grid2").Specific);
                oGrid2.DataTable = oDTRES;
                oGrid2.SelectionMode = BoMatrixSelect.ms_Single;
                oGrid2.Item.Enabled = false;


                /*oComboBox = (ComboBox)(oForm.Items.Item("Cliente").Specific);
                oForm.DataSources.UserDataSources.Add("Cliente", BoDataType.dt_SHORT_TEXT, 10);
                oComboBox.DataBind.SetBound(true, "", "Cliente");
                if (GlobalSettings.RunningUnderSQLServer)
                    s = @"SELECT 'Todos' Code, 'Todos' Name UNION ALL 
                          SELECT T1.FldValue Code, T1.Descr Name
                          FROM CUFD T0
                          JOIN UFD1 T1 ON T1.TableID = T0.TableID
                                      AND T1.FieldID = T0.FieldID
                         WHERE T0.TableID = '{0}'
                           AND T0.AliasID = 'EstadoC'";
                else
                    s = @"SELECT 'Todos' ""Code"", 'Todos' ""Name"" FROM DUMMY UNION ALL
                          SELECT T1.""FldValue"" ""Code"", T1.""Descr"" ""Name""
                          FROM ""CUFD"" T0
                          JOIN ""UFD1"" T1 ON T1.""TableID"" = T0.""TableID""
                                      AND T1.""FieldID"" = T0.""FieldID""
                         WHERE T0.""TableID"" = '{0}'
                           AND T0.""AliasID"" = 'EstadoC'";
                s = String.Format(s, "@VID_FEDTEVTA");
                oRecordSet.DoQuery(s);
                FSBOf.FillCombo(oComboBox, ref oRecordSet, false);
                oComboBox.Select("Todos", BoSearchKey.psk_ByValue);

                oComboBox = (ComboBox)(oForm.Items.Item("SII").Specific);
                oForm.DataSources.UserDataSources.Add("SII", BoDataType.dt_SHORT_TEXT, 10);
                oComboBox.DataBind.SetBound(true, "", "SII");
                if (GlobalSettings.RunningUnderSQLServer)
                    s = @"SELECT 'Todos' Code, 'Todos' Name UNION ALL
                          SELECT T1.FldValue Code, T1.Descr Name
                          FROM CUFD T0
                          JOIN UFD1 T1 ON T1.TableID = T0.TableID
                                      AND T1.FieldID = T0.FieldID
                         WHERE T0.TableID = '{0}'
                           AND T0.AliasID = 'EstadoSII'";
                else
                    s = @"SELECT 'Todos' ""Code"", 'Todos' ""Name"" FROM DUMMY UNION ALL
                          SELECT T1.""FldValue"" ""Code"", T1.""Descr"" ""Name""
                          FROM ""CUFD"" T0
                          JOIN ""UFD1"" T1 ON T1.""TableID"" = T0.""TableID""
                                      AND T1.""FieldID"" = T0.""FieldID""
                         WHERE T0.""TableID"" = '{0}'
                           AND T0.""AliasID"" = 'EstadoSII'";
                s = String.Format(s, "@VID_FEDTEVTA");
                oRecordSet.DoQuery(s);
                FSBOf.FillCombo(oComboBox, ref oRecordSet, false);
                oComboBox.Select("Todos", BoSearchKey.psk_ByValue);

                oComboBox = (ComboBox)(oForm.Items.Item("Ley").Specific);
                oForm.DataSources.UserDataSources.Add("Ley", BoDataType.dt_SHORT_TEXT, 10);
                oComboBox.DataBind.SetBound(true, "", "Ley");
                if (GlobalSettings.RunningUnderSQLServer)
                    s = @"SELECT 'Todos' Code, 'Todos' Name UNION ALL
                          SELECT T1.FldValue Code, T1.Descr Name
                          FROM CUFD T0
                          JOIN UFD1 T1 ON T1.TableID = T0.TableID
                                      AND T1.FieldID = T0.FieldID
                         WHERE T0.TableID = '{0}'
                           AND T0.AliasID = 'EstadoLey'";
                else
                    s = @"SELECT 'Todos' ""Code"", 'Todos' ""Name"" FROM DUMMY UNION ALL
                          SELECT T1.""FldValue"" ""Code"", T1.""Descr"" ""Name""
                          FROM ""CUFD"" T0
                          JOIN ""UFD1"" T1 ON T1.""TableID"" = T0.""TableID""
                                      AND T1.""FieldID"" = T0.""FieldID""
                         WHERE T0.""TableID"" = '{0}'
                           AND T0.""AliasID"" = 'EstadoLey'";
                s = String.Format(s, "@VID_FEDTEVTA");
                oRecordSet.DoQuery(s);
                FSBOf.FillCombo(oComboBox, ref oRecordSet, false);
                oComboBox.Select("Todos", BoSearchKey.psk_ByValue);
                */

                oForm.Items.Item("Chk_AC").Visible = false;
                oForm.Items.Item("Chk_RC").Visible = false;
                oForm.Items.Item("Chk_LN").Visible = false;
                oForm.Items.Item("lblAC").Visible = false;
                oForm.Items.Item("lblRC").Visible = false;


                //Cargar DT con Datos de Parametros
                CargarDatosParametros();
                BuscarDocumento();

                oForm.Items.Item("Folder1").Click(BoCellClickType.ct_Regular);
                oForm.PaneLevel = 2;
                oForm.PaneLevel = 1;
                oForm.Items.Item("Folder2").Enabled = false;

                string DocPrelim = oDTParams.GetValue("FProv", 0).ToString().Trim();
                oGrid2 = (Grid)(oForm.Items.Item("grid2").Specific);
                ((EditTextColumn)oGrid2.Columns.Item("Documento")).LinkedObjectType = DocPrelim == "Y" ? "112" : "18";

                oForm.Items.Item("lblTD").Visible = false;
                oForm.Items.Item("lblSL").Visible = false;
                oUD_LN = oForm.DataSources.UserDataSources.Item("UD_SEL");
                oUD_LN.ValueEx = "N";

            }
            catch (Exception e)
            {
                OutLog("InitForm: " + e.Message + " ** Trace: " + e.StackTrace);
                FSBOApp.MessageBox(e.Message + " ** Trace: " + e.StackTrace, 1, "Ok", "", "");
            }
            finally
            {
                if (oForm != null)
                    oForm.Freeze(false);
                oForm.Visible = true;
            }


            return Result;
        }//fin InitForm

        private void CargarDatosParametros()
        {
            try
            {
                if (GlobalSettings.RunningUnderSQLServer)
                    s = @"SELECT ISNULL(U_CrearDocC,'N')        'Crear'
                                , ISNULL(U_FProv,'Y')           'FProv' 
                                , ISNULL(U_DifPor,0)          'DifPor'
                                , ISNULL(U_DiasOC,0)          'DiasOC'
                                , ISNULL(U_DifMon,0)          'DifMon'
                                , ISNULL(U_TipoDif,'M')         'TipoDif'
                                , ISNULL(U_EntMer,'N')          'EntMer'
                                , ISNULL(U_CodEM,'N')           'CodEM'
                                , ISNULL(U_VListNegra,'N')      'VListNegra'
                                , ISNULL(U_MListNegra,'N')      'MListNegra'
                                , ISNULL(U_MListBlanca,'N')     'MListBlanca'
                                , ISNULL(U_MFacDifer,'N')       'MFacDifer'
                                , ISNULL(U_CEdoPortal,'Y')      'CEdoPortal'
                                , ISNULL(U_BuscarEMDocNum,'Y')  'BuscarEMDocNum'
                         FROM [@VID_FEPARAM]";
                else
                    s = @"SELECT IFNULL(""U_CrearDocC"",'N')    ""Crear""
                                , IFNULL(""U_FProv"",'Y')       ""FProv"" 
                                , IFNULL(""U_DiasOC"",0)      ""DiasOC""
                                , IFNULL(""U_DifPor"",0)      ""DifPor""
                                , IFNULL(""U_DifMon"",0)      ""DifMon""
                                , IFNULL(""U_TipoDif"",'M')     ""TipoDif""
                                , IFNULL(""U_EntMer"",'N')      ""EntMer""
                                , IFNULL(""U_CodEM"",'N')       ""CodEM""
                                , IFNULL(""U_VListNegra"",'N')  ""VListNegra""
                                , IFNULL(""U_MListNegra"",'N')  ""MListNegra""
                                , IFNULL(""U_MListBlanca"",'N') ""MListBlanca""
                                , IFNULL(""U_MFacDifer"",'N')   ""MFacDifer""
                                , IFNULL(""U_CEdoPortal"",'Y')  ""CEdoPortal""
                                , IFNULL(""U_BuscarEMDocNum"",'Y') ""BuscarEMDocNum""
                         FROM ""@VID_FEPARAM"" ";
                oDTParams.ExecuteQuery(s);
            }
            catch (Exception e)
            {
                OutLog("InitForm: " + e.Message + " ** Trace: " + e.StackTrace);
                FSBOApp.MessageBox(e.Message + " ** Trace: " + e.StackTrace, 1, "Ok", "", "");
            }
        }

        public new void FormEvent(String FormUID, ref SAPbouiCOM.ItemEvent pVal, ref Boolean BubbleEvent)
        {
            SAPbouiCOM.DataTable oDataTable;
            SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
            base.FormEvent(FormUID, ref pVal, ref BubbleEvent);

            try
            {
                if ((pVal.EventType == BoEventTypes.et_MATRIX_LINK_PRESSED) && (pVal.BeforeAction) && (pVal.ItemUID == "grid") && (pVal.ColUID == "OC" || pVal.ColUID == "EM"))
                {
                    oForm.Freeze(true);
                    oForm.Items.Item("Logo").Click();

                    if (pVal.ColUID == "OC")
                    {
                        try
                        {
                            if (GlobalSettings.RunningUnderSQLServer)
                                s = @"SELECT DocEntry FROM OPOR WHERE DocNum = {0}";
                            else
                                s = @"SELECT ""DocEntry"" FROM ""OPOR"" WHERE ""DocNum"" = {0}";

                            string sDocNum = oGrid.DataTable.GetValue("OC", pVal.Row).ToString().Trim();
                            s = string.Format(s, FSBOf.IsNumber(sDocNum) ? sDocNum : "0");

                            oRecordSet.DoQuery(s);

                            TempDocNumLink = sDocNum;
                            oGrid.DataTable.SetValue("OC", pVal.Row, oRecordSet.Fields.Item("DocEntry").Value.ToString());
                            if (oGrid.CommonSetting.GetCellEditable(pVal.Row, 17)) oGrid.SetCellFocus(pVal.Row, 17);
                        }
                        catch { }
                    }
                    else if (pVal.ColUID == "EM")
                    {
                        string sBuscarEMDocNum = oDTParams.GetValue("BuscarEMDocNum", 0).ToString().Trim();
                        if (sBuscarEMDocNum == "Y") //buscar EM por DocNum
                        {
                            try
                            {
                                if (GlobalSettings.RunningUnderSQLServer)
                                    s = @"SELECT DocEntry FROM OPDN WHERE DocNum = {0}";
                                else
                                    s = @"SELECT ""DocEntry"" FROM ""OPDN"" WHERE ""DocNum"" = {0}";

                                string sDocNum = oGrid.DataTable.GetValue("EM", pVal.Row).ToString().Trim();
                                s = string.Format(s, FSBOf.IsNumber(sDocNum) ? sDocNum : "0");

                                oRecordSet.DoQuery(s);

                                TempDocNumLink = sDocNum;
                                oGrid.DataTable.SetValue("EM", pVal.Row, oRecordSet.Fields.Item("DocEntry").Value.ToString());
                                if (oGrid.CommonSetting.GetCellEditable(pVal.Row, 19)) oGrid.SetCellFocus(pVal.Row, 19);
                            }
                            catch { }
                        }
                        else // Buscar EM Por FolioNum
                        {
                            try
                            {
                                if (GlobalSettings.RunningUnderSQLServer)
                                    s = @"SELECT DocEntry FROM OPDN WHERE FolioNum = {0}";
                                else
                                    s = @"SELECT ""DocEntry"" FROM ""OPDN"" WHERE ""FolioNum"" = {0}";

                                string sDocNum = oGrid.DataTable.GetValue("EM", pVal.Row).ToString().Trim();
                                s = string.Format(s, FSBOf.IsNumber(sDocNum) ? sDocNum : "0");

                                oRecordSet.DoQuery(s);

                                TempDocNumLink = sDocNum;
                                oGrid.DataTable.SetValue("EM", pVal.Row, oRecordSet.Fields.Item("DocEntry").Value.ToString());
                                if (oGrid.CommonSetting.GetCellEditable(pVal.Row, 19)) oGrid.SetCellFocus(pVal.Row, 19);
                            }
                            catch { }
                        }
                    }
                }

                if ((pVal.EventType == BoEventTypes.et_MATRIX_LINK_PRESSED) && (!pVal.BeforeAction) && (pVal.ItemUID == "grid") && (pVal.ColUID == "OC" || pVal.ColUID == "EM"))
                {
                    if (pVal.ColUID == "OC")
                    {
                        oGrid.DataTable.SetValue("OC", pVal.Row, TempDocNumLink);
                    }
                    else if (pVal.ColUID == "EM")
                    {
                        oGrid.DataTable.SetValue("EM", pVal.Row, TempDocNumLink);

                    }
                    oForm.Freeze(false);
                }

                if ((pVal.EventType == BoEventTypes.et_ITEM_PRESSED) && (!pVal.BeforeAction) && (pVal.ItemUID == "Buscar"))
                {
                    if (FSBOApp.MessageBox("Este Proceso Cargará de Nuevos los Datos, Se Perderan los Cambios no Registrados", 1, "Ok", "Cancelar", "") == 1)
                    {
                        oForm.Freeze(true);
                        CargarDatosParametros();
                        BuscarDocumento();
                        oForm.Freeze(false);
                    }
                }

                if ((pVal.EventType == BoEventTypes.et_ITEM_PRESSED) && (!pVal.BeforeAction) && (pVal.ItemUID == "Validar"))
                {
                    string sRUT = oGrid.DataTable.GetValue("U_RUT", 0).ToString().Trim();
                    if (oGrid.DataTable.Rows.Count > 0 && sRUT != "")
                    {
                        oForm.Freeze(true);
                        CargarDatosParametros();
                        ValidarDocumentos();
                        oForm.Freeze(false);
                    }
                }

                if ((pVal.EventType == BoEventTypes.et_ITEM_PRESSED) && (pVal.ItemUID == "Procesar") && (pVal.BeforeAction))
                {
                    BubbleEvent = false;

                    bool ValidaSelec = true;
                    for (Int32 i = 0; i < oGrid.DataTable.Rows.Count; i++)
                    {
                        var ocheckColumn = (CheckBoxColumn)oGrid.Columns.Item("Acepta");
                        string sEstado = oGrid.DataTable.GetValue("U_EstadoLey", i).ToString();

                        string ValidarEM = oDTParams.GetValue("EntMer", 0).ToString().Trim();
                        string sEMCode = oDTParams.GetValue("CodEM", 0).ToString().Trim();
                        string sDocOrig = ValidarEM == "Y" ? oGrid.DataTable.GetValue("EM", i).ToString().Trim() : oGrid.DataTable.GetValue("OC", i).ToString().Trim();
                        string sTipoDocOrig = ValidarEM == "Y" ? "Entrada de Mercancia" : "Orden de Compra";
                        string sCodDocOrig = ValidarEM == "Y" ? sEMCode : "OC";
                        string sFormPag = oGrid.DataTable.GetValue("U_FmaPago", i).ToString();
                        string sRutProv = oGrid.DataTable.GetValue("U_RUT", i).ToString().Trim();

                        if (sEstado.Trim() == "" && ocheckColumn.IsChecked(i) && sFormPag != "1" && sFormPag != "3")
                        {
                            ValidaSelec = false;
                            FSBOApp.MessageBox("Los Documentos Seleccionados deben tener asignado un Estado. Folio : " + oGrid.DataTable.GetValue("U_Folio", i).ToString(), 1, "Ok", "");
                            break;
                        }
                        else if (sDocOrig.Trim() == "" && ocheckColumn.IsChecked(i) && (sEstado.Trim() == "ACD" || sFormPag == "1" || sFormPag == "3"))
                        {
                            ValidaSelec = false;
                            FSBOApp.MessageBox("Los Documentos Seleccionados deben tener asignado una " + sTipoDocOrig + @" de Referencia. Folio : " + oGrid.DataTable.GetValue("U_Folio", i).ToString(), 1, "Ok", "");
                            break;
                        }
                        else if (ocheckColumn.IsChecked(i))
                        {
                            if (ValidaDocRefProveedor(sRutProv, sCodDocOrig, sDocOrig))
                            {
                                ValidaSelec = false;
                                FSBOApp.MessageBox("Para el Proveedor " + sRutProv + " , Folio : '" + oGrid.DataTable.GetValue("U_Folio", i).ToString() + "' Ya se Asigno la " + sTipoDocOrig + " Numero: '" + sDocOrig + @"' a una Compra Generada", 1, "Ok", "");
                                break;
                            }
                            string resp = ValidaDocEstado(sCodDocOrig, sDocOrig);
                            if (resp.Trim().Length > 0)
                            {
                                ValidaSelec = false;
                                FSBOApp.MessageBox("Para el Proveedor " + sRutProv + " , Folio : '" + oGrid.DataTable.GetValue("U_Folio", i).ToString() + "'  " + resp, 1, "Ok", "");
                                break;
                            }
                        }
                    }

                    if (ValidaSelec)
                    {
                        CargarDatosParametros();

                        if (ActualizarEstados())
                        {
                            oForm.Freeze(true);
                            BuscarDocumento(false);
                            oForm.Items.Item("Folder2").Enabled = true;
                            ((Grid)(oForm.Items.Item("grid2").Specific)).AutoResizeColumns();
                            oForm.Freeze(false);
                        }
                    }
                }

                if ((pVal.EventType == BoEventTypes.et_COMBO_SELECT) && (!pVal.BeforeAction) && (pVal.ColUID == "U_EstadoLey"))
                {
                    string sEstado = oGrid.DataTable.GetValue("U_EstadoLey", pVal.Row).ToString();
                    if (sEstado.Length > 0)
                    {
                        switch (sEstado)
                        {
                            case "ACD":
                                TotAcep += 1;
                                break;
                            case "ERM":
                                break;
                            default:
                                TotRecl += 1;
                                break;
                        }
                    }
                }

                if ((pVal.EventType == BoEventTypes.et_ITEM_PRESSED) && (pVal.BeforeAction) && (pVal.ColUID == "Acepta"))
                {
                    try
                    {
                        var ocheckColumn = (CheckBoxColumn)oGrid.Columns.Item("Acepta");
                        var ocbxColumn = (ComboBoxColumn)oGrid.Columns.Item("U_EstadoLey");
                        // string sEstado = ocheckColumn.IsChecked(pVal.Row) ? "ACD" : "RCD";
                        if (ocheckColumn.IsChecked(pVal.Row))
                            TotSele += 1;
                        else
                            TotSele -= 1;
                        ((StaticText)oForm.Items.Item("lblSL").Specific).Caption = "Total Seleccionados : " + TotSele.ToString();
                        //ocbxColumn.SetSelectedValue(pVal.Row, ocbxColumn.ValidValues.Item(ind));
                        //oGrid.DataTable.SetValue("U_EstadoLey", pVal.Row, sEstado);
                    }
                    catch { }
                }

                if ((pVal.EventType == BoEventTypes.et_DOUBLE_CLICK) && (!pVal.BeforeAction) && (pVal.ColUID == "U_Folio"))
                    MostrarPDF(pVal.Row);

                if ((pVal.EventType == BoEventTypes.et_FORM_RESIZE) && (!pVal.BeforeAction))
                {
                    oForm.Items.Item("Validar").Left = oForm.Width - 163;
                    oForm.Items.Item("Buscar").Left = oForm.Items.Item("Validar").Left + 71;
                }

                if ((pVal.EventType == BoEventTypes.et_CLICK) && (!pVal.BeforeAction) && (pVal.ItemUID == "grid2"))
                {
                    oGrid2.Rows.SelectedRows.Add(pVal.Row);
                }

                if ((pVal.EventType == BoEventTypes.et_DOUBLE_CLICK) && (!pVal.BeforeAction) && (pVal.ColUID == "Acepta") && (pVal.Row == -1))
                {
                    if (((CheckBox)oForm.Items.Item("chk_sel").Specific).Checked)
                    {
                        ((CheckBox)oForm.Items.Item("chk_sel").Specific).Checked = false;
                    }
                    else
                    {
                        ((CheckBox)oForm.Items.Item("chk_sel").Specific).Checked = true;
                    }

                }

                if (pVal.ItemUID == "chk_sel" && (pVal.EventType == BoEventTypes.et_ITEM_PRESSED) && (!pVal.BeforeAction))
                {
                    if (((CheckBox)oForm.Items.Item("chk_sel").Specific).Checked)
                        Seleccion(true);
                    else
                        Seleccion(false);
                }
            }
            catch (Exception e)
            {
                if (FCmpny.InTransaction)
                    FCmpny.EndTransaction(BoWfTransOpt.wf_RollBack);
                FSBOApp.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("FormEvent: " + e.Message + " ** Trace: " + e.StackTrace);
            }

        }//fin FormEvent


        public new void FormDataEvent(ref BusinessObjectInfo BusinessObjectInfo, ref Boolean BubbleEvent)
        {
            base.FormDataEvent(ref BusinessObjectInfo, ref BubbleEvent);

            try
            {

            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText("FormDataEvent: " + e.Message + " ** Trace: " + e.StackTrace, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("FormDataEvent: " + e.Message + " ** Trace: " + e.StackTrace);
            }
        }//fin FormDataEvent


        public new void MenuEvent(ref MenuEvent pVal, ref Boolean BubbleEvent)
        {
            //Int32 Entry;
            base.MenuEvent(ref pVal, ref BubbleEvent);
            try
            {
                //1281 Buscar; 
                //1282 Crear
                //1284 cancelar; 
                //1285 Restablecer; 
                //1286 Cerrar; 
                //1288 Registro siguiente;
                //1289 Registro anterior; 
                //1290 Primer Registro; 
                //1291 Ultimo Registro; 

            }
            catch (Exception e)
            {
                FSBOApp.MessageBox(e.Message + " ** Trace: " + e.StackTrace, 1, "Ok", "", "");
                OutLog("MenuEvent: " + e.Message + " ** Trace: " + e.StackTrace);
            }
        }//fin MenuEvent

        private void Seleccion(Boolean bvalor)
        {
            try
            {
                oForm.Freeze(true);

                Grid oGrid = (Grid)oForm.Items.Item("grid").Specific;
                SAPbouiCOM.DataTable DT_GRID = oForm.DataSources.DataTables.Item("dt");

                if (DT_GRID.Rows.Count > 50)
                    FSBOApp.StatusBar.SetText(bvalor ? "Seleccionando Registros" : "Borrando Seleccion Registros", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);


                for (Int32 i = 0; i <= DT_GRID.Rows.Count - 1; i++)
                {
                    if (oGrid.CommonSetting.GetCellEditable(i + 1, 1))
                        if (bvalor)
                            DT_GRID.SetValue("Acepta", i, "Y");
                        else
                            DT_GRID.SetValue("Acepta", i, "N");
                }

                if (DT_GRID.Rows.Count > 50)
                    FSBOApp.StatusBar.SetText("", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None);

            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("Seleccion: " + e.Message + " ** Trace: " + e.StackTrace);
            }
            oForm.Freeze(false);
        }

        private void BuscarDocumento(bool MostrarMensajes = true)
        {
            try
            {
                FSBOApp.StatusBar.SetText("Cargando Documentos, Por Favor Espere ... ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                string sEMCode = oDTParams.GetValue("CodEM", 0).ToString().Trim();
                string ValidarEM = oDTParams.GetValue("EntMer", 0).ToString().Trim();

                if (GlobalSettings.RunningUnderSQLServer)
                    s = @"SELECT 
                               'N' 'Acepta'
                              ,T0.U_RUT
                              ,T0.U_Razon
	                          ,T0.U_TipoDoc
	                          ,T0.U_Folio
                              ,T0.U_Validacion
                              ,CASE WHEN ISNULL(T4.U_FmaPago,'2')  = '1' THEN 'Contado' 
                                    WHEN ISNULL(T4.U_FmaPago,'2')  = '2' THEN 'Credito' 
                                    WHEN ISNULL(T4.U_FmaPago,'2')  = '3' THEN 'Ent. Gratuita' 
                                    END as 'U_FPagoD'
                              ,T0.U_FechaEmi
                              ,CAST(' ' AS VARCHAR(30)) 'Tiempo Rest.'
	                          ,CAST(REPLACE(CONVERT(CHAR(10), T0.U_FechaRecep, 102),'.','-') +'  '+ 
								    CASE WHEN LEN(T0.U_HoraRecep) = 4 THEN LEFT(CAST(T0.U_HoraRecep AS VARCHAR(10)),2) + ':' + RIGHT(CAST(T0.U_HoraRecep AS VARCHAR(10)),2) + ':00'
									 	    WHEN LEN(T0.U_HoraRecep) = 3 THEN '0' + LEFT(CAST(T0.U_HoraRecep AS VARCHAR(10)),1) + ':' + RIGHT(CAST(T0.U_HoraRecep AS VARCHAR(10)),2) + ':00'
									 	    WHEN LEN(T0.U_HoraRecep) = 2 THEN '00:'+ CAST(T0.U_HoraRecep AS VARCHAR(10)) + ':00'
										    WHEN LEN(T0.U_HoraRecep) = 1 THEN '00:0' + CAST(T0.U_HoraRecep AS VARCHAR(10)) + ':00'
										    ELSE '00:00:00'
								    END AS VARCHAR(50)) 'U_FechaRecep'
                              ,CAST(REPLACE(CONVERT(CHAR(10), T0.U_FechaRecep, 102),'.','-') +'T'+ 
								    CASE WHEN LEN(T0.U_HoraRecep) = 4 THEN LEFT(CAST(T0.U_HoraRecep AS VARCHAR(10)),2) + ':' + RIGHT(CAST(T0.U_HoraRecep AS VARCHAR(10)),2) + ':00'
									 	 WHEN LEN(T0.U_HoraRecep) = 3 THEN '0' + LEFT(CAST(T0.U_HoraRecep AS VARCHAR(10)),1) + ':' + RIGHT(CAST(T0.U_HoraRecep AS VARCHAR(10)),2) + ':00'
									 	 WHEN LEN(T0.U_HoraRecep) = 2 THEN '00:'+ CAST(T0.U_HoraRecep AS VARCHAR(10)) + ':00'
										 WHEN LEN(T0.U_HoraRecep) = 1 THEN '00:0' + CAST(T0.U_HoraRecep AS VARCHAR(10)) + ':00'
										 ELSE '00:00:00'
								    END AS DATETIME) as 'RecepPortal'
	                          ,T0.U_IVA
	                          ,T0.U_Monto
	                          ,T0.U_EstadoLey
                              ,T0.U_EstadoLey 'EstadoLeyOld'
                              ,T0.DocEntry
                              ,T0.U_Xml                              
                              ,ISNULL((SELECT TOP 1 A.U_FolioRef FROM [@VID_FEXMLCR] A WHERE A.Code = T0.DocEntry AND A.U_TpoDocRef = '801'),'') 'OCOri'
                              ,CAST(' ' AS VARCHAR(30)) 'OC'
                              ,ISNULL((SELECT TOP 1 A.U_FolioRef FROM [@VID_FEXMLCR] A WHERE A.Code = T0.DocEntry AND A.U_TpoDocRef = '" + sEMCode + @"'),'') 'EMOri'
                              ,CAST(' ' AS VARCHAR(30)) 'EM'
                              ,CASE WHEN T0.U_TipoDoc = '61' THEN ISNULL((SELECT TOP 1 A.U_FolioRef FROM [@VID_FEXMLCR] A WHERE A.Code = T0.DocEntry AND A.U_TpoDocRef = '33'),'') 
									WHEN T0.U_TipoDoc = '56' THEN ISNULL((SELECT TOP 1 A.U_FolioRef FROM [@VID_FEXMLCR] A WHERE A.Code = T0.DocEntry AND (A.U_TpoDocRef = '61' or A.U_TpoDocRef = '33')),'') 
									ELSE ''
							   END 'RefOri'
							 ,CAST(' ' AS VARCHAR(30)) 'Ref'
							 ,ISNULL((SELECT TOP 1 A.U_TpoDocRef FROM [@VID_FEXMLCR] A WHERE A.Code = T0.DocEntry),'') 'TpoRefOri'
							 ,CAST(' ' AS VARCHAR(30)) 'TpoRef'
                              ,CAST(' ' AS VARCHAR(250)) 'Desc_Valida'
                              ,ISNULL(T2.U_CardCode,CAST(' ' AS VARCHAR(30))) 'LB'
                              ,ISNULL(T3.U_CardCode,CAST(' ' AS VARCHAR(30))) 'LN'
                              ,CAST(' ' AS VARCHAR(30)) 'V_PORC'
                              ,CAST(' ' AS VARCHAR(30)) 'V_MONT'
                              ,CAST(' ' AS VARCHAR(30)) 'V_DIASOC'
                              ,CAST(' ' AS VARCHAR(30)) 'V_DIASEM'
                              ,CAST(' ' AS VARCHAR(30)) 'V_LB'
                              ,CAST(' ' AS VARCHAR(30)) 'V_LN'
                              ,CAST(' ' AS VARCHAR(30)) 'DiasRest'
                              ,CAST(' ' AS VARCHAR(30)) 'Color'
                              ,ISNULL(T1.CardCode,'') 'CardCode'
                              ,ISNULL(T4.U_FmaPago,'2') 'U_FmaPago'
                          FROM [@VID_FEDTECPRA] T0
                               LEFT JOIN OCRD T1 ON REPLACE(T1.LicTradNum,'.','') = T0.U_RUT AND T1.CardType = 'S' AND T1.validFor ='Y'
							   LEFT JOIN [@VID_FELISTABL] T2 ON T1.CardCode = T2.U_CardCode AND ISNULL(T2.U_Activado,'N') = 'Y'
							   LEFT JOIN [@VID_FELISTANE] T3 ON T1.CardCode = T3.U_CardCode AND ISNULL(T3.U_Activado,'N') = 'Y'
                               LEFT JOIN [@VID_FEXMLC] T4 ON T0.DocEntry = T4.Code
                         WHERE (ISNULL(T0.U_EstadoLey,'') = '' OR ISNULL(T0.U_EstadoLey,'') = 'ERM')
                           AND T0.U_TipoDoc IN ('33', '34', '61', '56')
                         ORDER BY 10 ASC, 2, 5 ";
                else
                    s = @"SELECT
                               'N' ""Acepta"" 
                              ,T0.""U_RUT""
                              ,T0.""U_Razon""
	                          ,T0.""U_TipoDoc""
	                          ,T0.""U_Folio""
                              ,T0.""U_Validacion""
                              ,CASE WHEN IFNULL(T4.""U_FmaPago"",'2')  = '1' THEN 'Contado' 
                                    WHEN IFNULL(T4.""U_FmaPago"",'2')  = '2' THEN 'Credito' 
                                    WHEN IFNULL(T4.""U_FmaPago"",'2')  = '3' THEN 'Ent. Gratuita' 
                                    END as ""U_FPagoD""
                              ,T0.""U_FechaEmi""
                              ,CAST(' ' AS VARCHAR(30)) ""Tiempo Rest.""
	                          ,CAST(TO_VARCHAR(T0.""U_FechaRecep"", 'yyyy-MM-dd') ||'T'|| 
								   CASE WHEN LENGTH(T0.""U_HoraRecep"") = 4 THEN LEFT(CAST(T0.""U_HoraRecep"" AS VARCHAR(10)),2) || ':' || RIGHT(CAST(T0.""U_HoraRecep"" AS VARCHAR(10)),2) || ':00'
										WHEN LENGTH(T0.""U_HoraRecep"") = 3 THEN '0' || LEFT(CAST(T0.""U_HoraRecep"" AS VARCHAR(10)),1) || ':' || RIGHT(CAST(T0.""U_HoraRecep"" AS VARCHAR(10)),2) || ':00'
										WHEN LENGTH(T0.""U_HoraRecep"") = 2 THEN '00:' || CAST(T0.""U_HoraRecep"" AS VARCHAR(10)) || ':00'
										WHEN LENGTH(T0.""U_HoraRecep"") = 1 THEN '00:0' || CAST(T0.""U_HoraRecep"" AS VARCHAR(10)) || ':00'
										ELSE '00:00:00'
								   END AS VARCHAR(50)) ""U_FechaRecep""
                              ,CAST(TO_VARCHAR(T0.""U_FechaRecep"", 'yyyy-MM-dd') ||'T'|| 
								   CASE WHEN LENGTH(T0.""U_HoraRecep"") = 4 THEN LEFT(CAST(T0.""U_HoraRecep"" AS VARCHAR(10)),2) || ':' || RIGHT(CAST(T0.""U_HoraRecep"" AS VARCHAR(10)),2) || ':00'
										WHEN LENGTH(T0.""U_HoraRecep"") = 3 THEN '0' || LEFT(CAST(T0.""U_HoraRecep"" AS VARCHAR(10)),1) || ':' || RIGHT(CAST(T0.""U_HoraRecep"" AS VARCHAR(10)),2) || ':00'
										WHEN LENGTH(T0.""U_HoraRecep"") = 2 THEN '00:' || CAST(T0.""U_HoraRecep"" AS VARCHAR(10)) || ':00'
										WHEN LENGTH(T0.""U_HoraRecep"") = 1 THEN '00:0' || CAST(T0.""U_HoraRecep"" AS VARCHAR(10)) || ':00'
										ELSE '00:00:00'
								   END AS DATETIME) as ""RecepPortal""
	                          ,T0.""U_IVA""
	                          ,T0.""U_Monto""
	                          ,T0.""U_EstadoLey""
                              ,T0.""U_EstadoLey"" ""EstadoLeyOld""
                              ,T0.""DocEntry""
                              ,T0.""U_Xml""
                              ,IFNULL((SELECT MAX(A.""U_FolioRef"") FROM ""@VID_FEXMLCR"" A WHERE A.""Code"" = T0.""DocEntry"" AND A.""U_TpoDocRef"" = '801'),'') ""OCOri""
                              ,CAST(' ' AS VARCHAR(30)) ""OC""
                              ,IFNULL((SELECT MAX(A.""U_FolioRef"") FROM ""@VID_FEXMLCR"" A WHERE A.""Code"" = T0.""DocEntry"" AND A.""U_TpoDocRef"" = '" + sEMCode + @"'),'') ""EMOri""
                              ,CAST(' ' AS VARCHAR(30)) ""EM""
                              ,CASE WHEN T0.""U_TipoDoc"" = '61' THEN IFNULL((SELECT MAX (A.""U_FolioRef"") FROM ""@VID_FEXMLCR"" A WHERE A.""Code"" = T0.""DocEntry"" AND A.""U_TpoDocRef"" = '33'),'') 
                                    WHEN T0.""U_TipoDoc"" = '56' THEN IFNULL((SELECT MAX (A.""U_FolioRef"") FROM ""@VID_FEXMLCR"" A WHERE A.""Code"" = T0.""DocEntry"" AND (A.""U_TpoDocRef"" = '61' OR A.""U_TpoDocRef"" = '33')),'') 
                                    ELSE ''
                                END AS ""RefOri""
                              ,CAST(' ' AS VARCHAR(30)) ""Ref""
                              ,IFNULL((SELECT MAX(A.""U_TpoDocRef"") FROM ""@VID_FEXMLCR"" A WHERE A.""Code"" = T0.""DocEntry"" ),'') ""TpoRefOri""
                              ,CAST(' ' AS VARCHAR(30)) ""TpoRef""
                              ,CAST(' ' AS VARCHAR(250)) ""Desc_Valida""
                              ,IFNULL(T2.""U_CardCode"", CAST(' ' AS VARCHAR(30))) ""LB""
                              ,IFNULL(T3.""U_CardCode"", CAST(' ' AS VARCHAR(30))) ""LN""
                              ,CAST(' ' AS VARCHAR(30)) ""V_PORC""
                              ,CAST(' ' AS VARCHAR(30)) ""V_MONT""
                              ,CAST(' ' AS VARCHAR(30)) ""V_DIASOC""
                              ,CAST(' ' AS VARCHAR(30)) ""V_DIASEM""
                              ,CAST(' ' AS VARCHAR(30)) ""V_LB""
                              ,CAST(' ' AS VARCHAR(30)) ""V_LN""
                              ,CAST(' ' AS VARCHAR(30)) ""DiasRest""
                              ,CAST(' ' AS VARCHAR(30)) ""Color""
                              ,IFNULL(T1.""CardCode"",'') ""CardCode""
                              ,IFNULL(T4.""U_FmaPago"",'2') ""U_FmaPago""
                          FROM ""@VID_FEDTECPRA"" T0
                               LEFT JOIN ""OCRD"" T1 ON REPLACE(T1.""LicTradNum"",'.','') = T0.""U_RUT"" AND T1.""CardType"" = 'S' AND T1.""validFor"" ='Y'
							   LEFT JOIN ""@VID_FELISTABL"" T2 ON T1.""CardCode"" = T2.""U_CardCode"" AND IFNULL(T2.""U_Activado"",'N') = 'Y'
							   LEFT JOIN ""@VID_FELISTANE"" T3 ON T1.""CardCode"" = T3.""U_CardCode"" AND IFNULL(T3.""U_Activado"",'N') = 'Y'
                               LEFT JOIN ""@VID_FEXMLC"" T4 ON T0.""DocEntry"" = T4.""Code""
                         WHERE (IFNULL(T0.""U_EstadoLey"",'') = '' OR IFNULL(T0.""U_EstadoLey"",'') = 'ERM')
                           AND T0.""U_TipoDoc"" IN ('33', '34', '61' , '56')
                         ORDER BY 10 ASC, 2, 5 ";

                oGrid = (Grid)(oForm.Items.Item("grid").Specific);
                oGrid.DataTable.ExecuteQuery(s);

                oGrid.Columns.Item("Acepta").Type = BoGridColumnType.gct_CheckBox;
                var ocheckColumns = (GridColumn)(oGrid.Columns.Item("Acepta"));
                var ocheckColumn = (CheckBoxColumn)(ocheckColumns);
                ocheckColumn.Editable = true;
                ocheckColumn.TitleObject.Caption = "Selec.";

                oGrid.Columns.Item("U_RUT").Type = BoGridColumnType.gct_EditText;
                var oColumn = (GridColumn)(oGrid.Columns.Item("U_RUT"));
                var oEditColumn = (EditTextColumn)(oColumn);
                oEditColumn.Editable = false;
                oEditColumn.TitleObject.Caption = "RUT";

                oGrid.Columns.Item("U_Razon").Type = BoGridColumnType.gct_EditText;
                oColumn = (GridColumn)(oGrid.Columns.Item("U_Razon"));
                oEditColumn = (EditTextColumn)(oColumn);
                oEditColumn.Editable = false;
                oEditColumn.TitleObject.Caption = "Razón Social";
                oEditColumn.Width = 220;

                oGrid.Columns.Item("U_TipoDoc").Type = BoGridColumnType.gct_EditText;
                oColumn = (GridColumn)(oGrid.Columns.Item("U_TipoDoc"));
                oEditColumn = (EditTextColumn)(oColumn);
                oEditColumn.Editable = false;
                oEditColumn.RightJustified = true;
                oEditColumn.TitleObject.Caption = "Tipo Doc";

                oGrid.Columns.Item("U_Folio").Type = BoGridColumnType.gct_EditText;
                oColumn = (GridColumn)(oGrid.Columns.Item("U_Folio"));
                oEditColumn = (EditTextColumn)(oColumn);
                oEditColumn.Editable = false;
                oEditColumn.RightJustified = true;
                oEditColumn.TitleObject.Caption = "Folio";

                oGrid.Columns.Item("U_Validacion").Type = BoGridColumnType.gct_EditText;
                oColumn = (GridColumn)(oGrid.Columns.Item("U_Validacion"));
                oEditColumn = (EditTextColumn)(oColumn);
                oEditColumn.Editable = false;
                oEditColumn.Visible = true;
                oEditColumn.TitleObject.Caption = "Indicador";

                oGrid.Columns.Item("U_FPagoD").Type = BoGridColumnType.gct_EditText;
                oColumn = (GridColumn)(oGrid.Columns.Item("U_FPagoD"));
                oEditColumn = (EditTextColumn)(oColumn);
                oEditColumn.Editable = false;
                oEditColumn.Visible = true;
                oEditColumn.TitleObject.Caption = "Forma Pago";

                oGrid.Columns.Item("U_FechaEmi").Type = BoGridColumnType.gct_EditText;
                oColumn = (GridColumn)(oGrid.Columns.Item("U_FechaEmi"));
                oEditColumn = (EditTextColumn)(oColumn);
                oEditColumn.Editable = false;
                oEditColumn.TitleObject.Caption = "Fecha Emisión";

                oGrid.Columns.Item("Tiempo Rest.").Type = BoGridColumnType.gct_EditText;
                oColumn = (GridColumn)(oGrid.Columns.Item("Tiempo Rest."));
                oEditColumn = (EditTextColumn)(oColumn);
                oEditColumn.Editable = false;
                oEditColumn.TitleObject.Caption = "Tiempo Rest.";

                oGrid.Columns.Item("U_FechaRecep").Type = BoGridColumnType.gct_EditText;
                oColumn = (GridColumn)(oGrid.Columns.Item("U_FechaRecep"));
                oEditColumn = (EditTextColumn)(oColumn);
                oEditColumn.Editable = false;
                oEditColumn.TitleObject.Caption = "Fecha Recepción";

                oGrid.Columns.Item("RecepPortal").Visible = false;

                oGrid.Columns.Item("U_IVA").Type = BoGridColumnType.gct_EditText;
                oColumn = (GridColumn)(oGrid.Columns.Item("U_IVA"));
                oEditColumn = (EditTextColumn)(oColumn);
                oEditColumn.Editable = false;
                oEditColumn.RightJustified = true;
                oEditColumn.TitleObject.Caption = "IVA";

                oGrid.Columns.Item("U_Monto").Type = BoGridColumnType.gct_EditText;
                oColumn = (GridColumn)(oGrid.Columns.Item("U_Monto"));
                oEditColumn = (EditTextColumn)(oColumn);
                oEditColumn.Editable = false;
                oEditColumn.RightJustified = true;
                oEditColumn.TitleObject.Caption = "Monto";

                oGrid.Columns.Item("U_EstadoLey").Type = BoGridColumnType.gct_ComboBox;
                var ocbxColumns = (GridColumn)(oGrid.Columns.Item("U_EstadoLey"));
                var ocbxColumn = (ComboBoxColumn)(ocbxColumns);
                ocbxColumn.Editable = true;
                ocbxColumn.DisplayType = BoComboDisplayType.cdt_Description;
                ocbxColumn.TitleObject.Caption = "Estado Ley 20.956";
                if (GlobalSettings.RunningUnderSQLServer)
                    s = @"SELECT T1.FldValue Code, T1.Descr Name
                          FROM CUFD T0
                          JOIN UFD1 T1 ON T1.TableID = T0.TableID
                                      AND T1.FieldID = T0.FieldID
                         WHERE T0.TableID = '@VID_FEDTEVTA'
                           AND T0.AliasID = 'EstadoLey'
                           AND T1.FldValue <> 'ACO'";
                else
                    s = @"SELECT T1.""FldValue"" ""Code"", T1.""Descr"" ""Name""
                          FROM ""CUFD"" T0
                          JOIN ""UFD1"" T1 ON T1.""TableID"" = T0.""TableID""
                                      AND T1.""FieldID"" = T0.""FieldID""
                         WHERE T0.""TableID"" = '@VID_FEDTEVTA'
                           AND T0.""AliasID"" = 'EstadoLey'
                           AND T1.""FldValue"" <> 'ACO'";
                oRecordSet.DoQuery(s);
                FSBOf.FillComboGrid(ocbxColumns, ref oRecordSet, true);

                oGrid.Columns.Item("EstadoLeyOld").Type = BoGridColumnType.gct_EditText;
                oColumn = (GridColumn)(oGrid.Columns.Item("EstadoLeyOld"));
                oEditColumn = (EditTextColumn)(oColumn);
                oEditColumn.Editable = false;
                oEditColumn.Visible = false;

                oGrid.Columns.Item("DocEntry").Type = BoGridColumnType.gct_EditText;
                oColumn = (GridColumn)(oGrid.Columns.Item("DocEntry"));
                oEditColumn = (EditTextColumn)(oColumn);
                oEditColumn.Visible = false;

                oGrid.Columns.Item("U_Xml").Type = BoGridColumnType.gct_EditText;
                oEditColumn = ((EditTextColumn)oGrid.Columns.Item("U_Xml"));
                oEditColumn.Visible = false;

                oGrid.Columns.Item("OC").Type = BoGridColumnType.gct_EditText;
                oEditColumn = ((EditTextColumn)oGrid.Columns.Item("OC"));
                oEditColumn.Visible = true;
                oEditColumn.TitleObject.Caption = "Orden Compra";
                oEditColumn.Editable = ValidarEM == "Y" ? false : true;
                oEditColumn.LinkedObjectType = "22";

                oGrid.Columns.Item("OCOri").Type = BoGridColumnType.gct_EditText;
                oEditColumn = ((EditTextColumn)oGrid.Columns.Item("OCOri"));
                oEditColumn.Visible = false;  //#

                oGrid.Columns.Item("EM").Type = BoGridColumnType.gct_EditText;
                oEditColumn = ((EditTextColumn)oGrid.Columns.Item("EM"));
                oEditColumn.Visible = true;
                oEditColumn.TitleObject.Caption = "Entrada Mercancia";
                oEditColumn.Editable = ValidarEM == "Y" ? true : false;
                oEditColumn.LinkedObjectType = "20";

                oGrid.Columns.Item("Ref").Type = BoGridColumnType.gct_EditText;
                oEditColumn = ((EditTextColumn)oGrid.Columns.Item("Ref"));
                oEditColumn.Visible = false;

                oGrid.Columns.Item("RefOri").Type = BoGridColumnType.gct_EditText;
                oEditColumn = ((EditTextColumn)oGrid.Columns.Item("RefOri"));
                oEditColumn.Visible = true;
                oEditColumn.TitleObject.Caption = "Referencia";
                oEditColumn.Editable = false;
                oEditColumn.LinkedObjectType = "20"; // aca sacaquer consulta para obtener objecto

                oGrid.Columns.Item("TpoRefOri").Type = BoGridColumnType.gct_EditText;
                oEditColumn = ((EditTextColumn)oGrid.Columns.Item("TpoRefOri"));
                oEditColumn.Visible = true;
                oEditColumn.TitleObject.Caption = "Tipo Ref";
                oEditColumn.Editable = false;

                oGrid.Columns.Item("EMOri").Visible = false;//#

                oGrid.Columns.Item("Desc_Valida").Type = BoGridColumnType.gct_EditText;
                oEditColumn = ((EditTextColumn)oGrid.Columns.Item("Desc_Valida"));
                oEditColumn.Visible = true;
                oEditColumn.Editable = false;
                oEditColumn.TitleObject.Caption = "Validaciones";

                oGrid.Columns.Item("TpoRef").Visible = false;
                oGrid.Columns.Item("LB").Visible = false;
                oGrid.Columns.Item("LN").Visible = false;
                oGrid.Columns.Item("V_PORC").Visible = false;
                oGrid.Columns.Item("V_MONT").Visible = false;
                oGrid.Columns.Item("V_DIASOC").Visible = false;
                oGrid.Columns.Item("V_DIASEM").Visible = false;
                oGrid.Columns.Item("V_LB").Visible = false;
                oGrid.Columns.Item("V_LN").Visible = false;
                oGrid.Columns.Item("DiasRest").Visible = false;
                oGrid.Columns.Item("Color").Visible = false;
                oGrid.Columns.Item("CardCode").Visible = false;
                oGrid.Columns.Item("U_FmaPago").Visible = false;

                string sRUT = oGrid.DataTable.GetValue("U_RUT", 0).ToString().Trim();
                if (oGrid.DataTable.Rows.Count > 0 && sRUT != "")
                    ValidarDocumentos(MostrarMensajes);

                oGrid.AutoResizeColumns();

            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText("BuscarDocumento: " + e.Message + " ** Trace: " + e.StackTrace, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("BuscarDocumento: " + e.Message + " ** Trace: " + e.StackTrace);
            }
            finally
            {
                FSBOApp.StatusBar.SetText("", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None);
            }
        }

        private Boolean ActualizarEstados()
        {
            String URL = "http://portal.easydoc.cl/Consulta/GeneracionDte.aspx?RUT={0}&amp;FOLIO={1}&amp;TIPODTE={2}&amp;RUT1={3}&amp;CODAR={4}&amp;OP=31";
            String URLFinal;
            String TaxIdNum;
            String UserWS = "";
            String PassWS = "";
            WebRequest request;
            string postData;
            byte[] byteArray;
            Stream dataStream;
            WebResponse response;
            StreamReader reader;
            string responseFromServer;
            String EstadoOriginal;
            String EstadoFinal;
            String EstadoOld;
            Int32 lRetCode;
            SAPbouiCOM.Conditions oConditions;
            SAPbouiCOM.Condition oCondition;
            String EstadoDescrip;
            String ValidacionAct;
            String sFormPag;

            try
            {

                if (GlobalSettings.RunningUnderSQLServer)
                    s = @"SELECT ISNULL(TaxIdNum,'') TaxIdNum, CompnyName FROM OADM ";
                else
                    s = @"SELECT IFNULL(""TaxIdNum"",'') ""TaxIdNum"", ""CompnyName"" FROM ""OADM"" ";
                oRecordSet.DoQuery(s);
                if (oRecordSet.RecordCount == 0)
                {
                    FSBOApp.StatusBar.SetText("Debe ingresar RUT de Emisor, Gestión -> Inicialización Sistema -> Detalle Sociedad -> Datos de Contabilidad -> ID fiscal general 1", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else
                    TaxIdNum = ((System.String)oRecordSet.Fields.Item("TaxIdNum").Value).Trim();

                if (GlobalSettings.RunningUnderSQLServer)
                    s = @"SELECT ISNULL(U_UserWSCL,'') 'UserWS', ISNULL(U_PassWSCL,'') 'PassWS' FROM [@VID_FEPARAM]";
                else
                    s = @"SELECT IFNULL(""U_UserWSCL"",'') ""UserWS"", IFNULL(""U_PassWSCL"",'') ""PassWS"" FROM ""@VID_FEPARAM"" ";
                oRecordSet.DoQuery(s);
                if (((System.String)oRecordSet.Fields.Item("UserWS").Value).Trim() != "")
                    UserWS = Funciones.DesEncriptar(((System.String)oRecordSet.Fields.Item("UserWS").Value).Trim());
                if (((System.String)oRecordSet.Fields.Item("PassWS").Value).Trim() != "")
                    PassWS = Funciones.DesEncriptar(((System.String)oRecordSet.Fields.Item("PassWS").Value).Trim());

                oGrid = ((Grid)oForm.Items.Item("grid").Specific);

                for (Int32 i = 0; i < oGrid.DataTable.Rows.Count; i++)
                {
                    var ocheckColumn = (CheckBoxColumn)oGrid.Columns.Item("Acepta");

                    sFormPag = oGrid.DataTable.GetValue("U_FmaPago", i).ToString();

                    if (sFormPag == "1" || sFormPag == "3")// Si Forma Pago es Contado o Efectivo Aceptar
                        oGrid.DataTable.SetValue("U_EstadoLey", i, "ACD");

                    if (((System.String)oGrid.DataTable.GetValue("U_EstadoLey", i)).Trim() != "" && ocheckColumn.IsChecked(i))
                    {
                        string sEnviarEstadoPortal = oDTParams.GetValue("CEdoPortal", 0).ToString().Trim();
                        string CrearBaseEM = oDTParams.GetValue("EntMer", 0).ToString().Trim();
                        string sEMCode = oDTParams.GetValue("CodEM", 0).ToString().Trim();
                        string sTipoBase = CrearBaseEM == "Y" ? sEMCode : "OC";

                        string DocRefAct = CrearBaseEM == "Y" ? oGrid.DataTable.GetValue("EM", i).ToString() : oGrid.DataTable.GetValue("OC", i).ToString();

                        EstadoOld = ((System.String)oGrid.DataTable.GetValue("EstadoLeyOld", i)).Trim();
                        EstadoOriginal = ((System.String)oGrid.DataTable.GetValue("U_EstadoLey", i)).Trim();
                        if (EstadoOriginal == EstadoOld)
                            continue;
                        if ((EstadoOriginal == "ACD") && (EstadoOld != "ERM"))
                            EstadoFinal = "ERM";
                        else
                            EstadoFinal = EstadoOriginal;

                        if (GlobalSettings.RunningUnderSQLServer)
                            s = @"SELECT T1.FldValue Code, T1.Descr Name
                                  FROM CUFD T0
                                  JOIN UFD1 T1 ON T1.TableID = T0.TableID
                                              AND T1.FieldID = T0.FieldID
                                 WHERE T0.TableID = '@VID_FEDTEVTA'
                                   AND T0.AliasID = 'EstadoLey'
                                   AND T1.FldValue = '{0}'";
                        else
                            s = @"SELECT T1.""FldValue"" ""Code"", T1.""Descr"" ""Name""
                                  FROM ""CUFD"" T0
                                  JOIN ""UFD1"" T1 ON T1.""TableID"" = T0.""TableID""
                                              AND T1.""FieldID"" = T0.""FieldID""
                                 WHERE T0.""TableID"" = '@VID_FEDTEVTA'
                                   AND T0.""AliasID"" = 'EstadoLey'
                                   AND T1.""FldValue"" = '{0}'";
                        s = String.Format(s, EstadoFinal);
                        oRecordSet.DoQuery(s);
                        EstadoDescrip = ((System.String)oRecordSet.Fields.Item("Name").Value).Trim();
                        ValidacionAct = ((System.String)oGrid.DataTable.GetValue("Desc_Valida", i)).Trim();

                        URLFinal = URL.Replace("{0}", TaxIdNum.Replace("-", "").Replace(".", ""));
                        URLFinal = URLFinal.Replace("{1}", ((System.Int32)oGrid.DataTable.GetValue("U_Folio", i)).ToString());
                        URLFinal = URLFinal.Replace("{2}", ((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim());
                        URLFinal = URLFinal.Replace("{3}", ((System.String)oGrid.DataTable.GetValue("U_RUT", i)).Replace(".", "").Trim());
                        URLFinal = URLFinal.Replace("{4}", EstadoFinal);
                        URLFinal = URLFinal.Replace("&amp;", "&");

                        bool StatusResult = false;
                        string sDescripcion = "";
                        if (sFormPag != "1" && sFormPag != "3" && sEnviarEstadoPortal == "Y")
                        {
                            request = WebRequest.Create(URLFinal);
                            if ((UserWS != "") && (PassWS != ""))
                                request.Credentials = new NetworkCredential(UserWS, PassWS);
                            request.Method = "POST";
                            postData = "";//** xmlDOC.InnerXml;
                            byteArray = Encoding.UTF8.GetBytes(postData);
                            request.ContentType = "text/xml";
                            request.ContentLength = byteArray.Length;
                            dataStream = request.GetRequestStream();
                            dataStream.Write(byteArray, 0, byteArray.Length);
                            dataStream.Close();
                            response = request.GetResponse();
                            Console.WriteLine(((HttpWebResponse)(response)).StatusDescription);
                            dataStream = response.GetResponseStream();
                            reader = new StreamReader(dataStream);
                            responseFromServer = reader.ReadToEnd();
                            reader.Close();
                            dataStream.Close();
                            response.Close();
                            s = responseFromServer;
                            var results = JsonConvert.DeserializeObject<dynamic>(s);
                            var jStatus = results.Status;
                            var jCodigo = results.Codigo;
                            var jDescripcion = results.Descripcion;
                            sDescripcion = ((System.String)jDescripcion.Value).Trim();

                            request = null;
                            response = null;
                            dataStream = null;
                            reader = null;
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            StatusResult = ((System.String)jStatus.Value).Trim() == "OK" ? true : false;
                        }
                        else
                        {
                            StatusResult = true; //asumo que le envie al portal
                            sDescripcion = "";
                            if (sFormPag != "1" && sFormPag != "3" && sEnviarEstadoPortal == "N")
                            {
                                if ((EstadoFinal == "ACD") && ((((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim() == "33") || (((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim() == "34")))
                                    CrearDocto(((System.Int32)oGrid.DataTable.GetValue("DocEntry", i)), ((System.String)oGrid.DataTable.GetValue("U_RUT", i)), ((System.Int32)oGrid.DataTable.GetValue("U_Folio", i)), ((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim(), i);
                            }
                        }

                        if (StatusResult)
                        {
                            oDBDSHC.Clear();
                            oConditions = new SAPbouiCOM.Conditions();
                            oCondition = oConditions.Add();
                            oCondition.Alias = "DocEntry";
                            oCondition.Operation = BoConditionOperation.co_EQUAL;
                            var DocEntry = ((System.Int32)oGrid.DataTable.GetValue("DocEntry", i)).ToString();
                            oCondition.CondVal = DocEntry;
                            oDBDSHC.Query(oConditions);

                            oDBDSHC.SetValue("U_EstadoLey", 0, EstadoFinal);
                            oDBDSHC.SetValue("U_EstadoSII", 0, "A");
                            oDBDSHC.SetValue("U_Descrip", 0, EstadoDescrip);
                            oDBDSHC.SetValue("U_Validacion", 0, ValidacionAct);
                            oDBDSHC.SetValue("U_FechaMov", 0, DateTime.Now.Date.ToString("yyyyMMdd"));
                            oDBDSHC.SetValue("U_HoraMov", 0, DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString("00"));
                            oDBDSHC.SetValue("U_CodRefGen", 0, sTipoBase);
                            oDBDSHC.SetValue("U_FolioRefGen", 0, DocRefAct);

                            lRetCode = Funciones.UpdDataSourceInt1("VID_FEDTECPRA", oDBDSHC, "", null, "", null, "", null);
                            if (lRetCode == 0)
                            {
                                FSBOApp.StatusBar.SetText("No se actualizado tabla @VID_FEDTECPRA, dejar en estado " + EstadoFinal + ", RUT " + ((System.String)oGrid.DataTable.GetValue("U_RUT", i)).Trim() + ", Tipo doc " + ((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim() + ", Folio " + ((System.Int32)oGrid.DataTable.GetValue("U_Folio", i)).ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                OutLog("No se actualizado tabla @VID_FEDTECPRA, dejar en estado " + EstadoFinal + ", RUT " + ((System.String)oGrid.DataTable.GetValue("U_RUT", i)).Trim() + ", Tipo doc " + ((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim() + ", Folio " + ((System.Int32)oGrid.DataTable.GetValue("U_Folio", i)).ToString());
                            }
                            else
                            {
                                if (sFormPag != "1" && sFormPag != "3")
                                    FSBOApp.StatusBar.SetText("Documento actualizado en el portal, dejar en estado " + EstadoFinal + ", RUT " + ((System.String)oGrid.DataTable.GetValue("U_RUT", i)).Trim() + ", Tipo doc " + ((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim() + ", Folio " + ((System.Int32)oGrid.DataTable.GetValue("U_Folio", i)).ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                                if ((EstadoFinal == "ACD") && sEnviarEstadoPortal == "Y" && ((((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim() == "33") || (((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim() == "34")))
                                    CrearDocto(((System.Int32)oGrid.DataTable.GetValue("DocEntry", i)), ((System.String)oGrid.DataTable.GetValue("U_RUT", i)), ((System.Int32)oGrid.DataTable.GetValue("U_Folio", i)), ((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim(), i);
                            }

                            //para la aceptacion primero envia el recibo de mercaderia y luego debe enviar la aceptacion
                            if ((EstadoOriginal == "ACD") && (EstadoOld != "ERM"))
                            {
                                URLFinal = URL.Replace("{0}", TaxIdNum.Replace("-", "").Replace(".", ""));
                                URLFinal = URLFinal.Replace("{1}", ((System.Int32)oGrid.DataTable.GetValue("U_Folio", i)).ToString());
                                URLFinal = URLFinal.Replace("{2}", ((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim());
                                URLFinal = URLFinal.Replace("{3}", ((System.String)oGrid.DataTable.GetValue("U_RUT", i)).Replace(".", "").Trim());
                                URLFinal = URLFinal.Replace("{4}", EstadoOriginal);
                                URLFinal = URLFinal.Replace("&amp;", "&");

  
                                StatusResult = false;
                                sDescripcion = "";
                                if (sFormPag != "1" && sFormPag != "3" && sEnviarEstadoPortal == "Y")
                                {
                                    request = WebRequest.Create(URLFinal);
                                    if ((UserWS != "") && (PassWS != ""))
                                        request.Credentials = new NetworkCredential(UserWS, PassWS);
                                    request.Method = "POST";
                                    postData = "";//** xmlDOC.InnerXml;
                                    byteArray = Encoding.UTF8.GetBytes(postData);
                                    request.ContentType = "text/xml";
                                    request.ContentLength = byteArray.Length;
                                    dataStream = request.GetRequestStream();
                                    dataStream.Write(byteArray, 0, byteArray.Length);
                                    dataStream.Close();
                                    response = request.GetResponse();
                                    Console.WriteLine(((HttpWebResponse)(response)).StatusDescription);
                                    dataStream = response.GetResponseStream();
                                    reader = new StreamReader(dataStream);
                                    responseFromServer = reader.ReadToEnd();
                                    reader.Close();
                                    dataStream.Close();
                                    response.Close();
                                    s = responseFromServer;
                                    var results1 = JsonConvert.DeserializeObject<dynamic>(s);
                                    var jStatus1 = results1.Status;
                                    var jCodigo1 = results1.Codigo;
                                    var jDescripcion1 = results1.Descripcion;
                                    sDescripcion = ((System.String)jDescripcion1.Value).Trim();

                                    request = null;
                                    response = null;
                                    dataStream = null;
                                    reader = null;
                                    GC.Collect();
                                    GC.WaitForPendingFinalizers();
                                    StatusResult = ((System.String)jStatus1.Value).Trim() == "OK" ? true : false;
                                }
                                else
                                {
                                    StatusResult = true;
                                    if (sFormPag != "1" && sFormPag != "3" && sEnviarEstadoPortal == "N")
                                    {
                                        if ((((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim() == "33") || (((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim() == "34"))
                                            CrearDocto(((System.Int32)oGrid.DataTable.GetValue("DocEntry", i)), ((System.String)oGrid.DataTable.GetValue("U_RUT", i)), ((System.Int32)oGrid.DataTable.GetValue("U_Folio", i)), ((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim(), i);
                                    }

                                }
                                if (StatusResult)
                                {
                                    if (GlobalSettings.RunningUnderSQLServer)
                                        s = @"SELECT T1.FldValue Code, T1.Descr Name
                                              FROM CUFD T0
                                              JOIN UFD1 T1 ON T1.TableID = T0.TableID
                                                          AND T1.FieldID = T0.FieldID
                                             WHERE T0.TableID = '@VID_FEDTEVTA'
                                               AND T0.AliasID = 'EstadoLey'
                                               AND T1.FldValue = '{0}'";
                                    else
                                        s = @"SELECT T1.""FldValue"" ""Code"", T1.""Descr"" ""Name""
                                              FROM ""CUFD"" T0
                                              JOIN ""UFD1"" T1 ON T1.""TableID"" = T0.""TableID""
                                                          AND T1.""FieldID"" = T0.""FieldID""
                                             WHERE T0.""TableID"" = '@VID_FEDTEVTA'
                                               AND T0.""AliasID"" = 'EstadoLey'
                                               AND T1.""FldValue"" = '{0}'";
                                    s = String.Format(s, EstadoOriginal);
                                    oRecordSet.DoQuery(s);
                                    EstadoDescrip = ((System.String)oRecordSet.Fields.Item("Name").Value).Trim();

                                    oDBDSHC.Clear();
                                    oConditions = new SAPbouiCOM.Conditions();
                                    oCondition = oConditions.Add();
                                    oCondition.Alias = "DocEntry";
                                    oCondition.Operation = BoConditionOperation.co_EQUAL;
                                    oCondition.CondVal = DocEntry;
                                    oDBDSHC.Query(oConditions);

                                    oDBDSHC.SetValue("U_EstadoLey", 0, EstadoOriginal);
                                    oDBDSHC.SetValue("U_EstadoSII", 0, "A");
                                    oDBDSHC.SetValue("U_Descrip", 0, EstadoDescrip);
                                    oDBDSHC.SetValue("U_Validacion", 0, ValidacionAct);
                                    oDBDSHC.SetValue("U_FechaMov", 0, DateTime.Now.Date.ToString("yyyyMMdd"));
                                    oDBDSHC.SetValue("U_HoraMov", 0, DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString("00"));
                                    oDBDSHC.SetValue("U_CodRefGen", 0, sTipoBase);
                                    oDBDSHC.SetValue("U_FolioRefGen", 0, DocRefAct);

                                    if (Funciones.UpdDataSourceInt1("VID_FEDTECPRA", oDBDSHC, "", null, "", null, "", null) == 0)
                                    {
                                        FSBOApp.StatusBar.SetText("No se actualizado tabla @VID_FEDTECPRA, dejar en estado " + EstadoOriginal + ", RUT " + ((System.String)oGrid.DataTable.GetValue("U_RUT", i)).Trim() + ", Tipo doc " + ((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim() + ", Folio " + ((System.Int32)oGrid.DataTable.GetValue("U_Folio", i)).ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        OutLog("No se actualizado tabla @VID_FEDTECPRA, dejar en estado " + EstadoOriginal + ", RUT " + ((System.String)oGrid.DataTable.GetValue("U_RUT", i)).Trim() + ", Tipo doc " + ((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim() + ", Folio " + ((System.Int32)oGrid.DataTable.GetValue("U_Folio", i)).ToString());
                                    }
                                    else
                                    {
                                        if (sFormPag != "1" && sFormPag != "3")
                                            FSBOApp.StatusBar.SetText("Documento actualizado en el portal, dejar en estado " + EstadoOriginal + ", RUT " + ((System.String)oGrid.DataTable.GetValue("U_RUT", i)).Trim() + ", Tipo doc " + ((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim() + ", Folio " + ((System.Int32)oGrid.DataTable.GetValue("U_Folio", i)).ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                        if (sEnviarEstadoPortal == "Y")
                                        {
                                            if ((((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim() == "33") || (((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim() == "34"))
                                                CrearDocto(((System.Int32)oGrid.DataTable.GetValue("DocEntry", i)), ((System.String)oGrid.DataTable.GetValue("U_RUT", i)), ((System.Int32)oGrid.DataTable.GetValue("U_Folio", i)), ((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim(), i);
                                        }
                                    }
                                }
                                else
                                {
                                    FSBOApp.StatusBar.SetText("No se actualizado en el portal(" + sDescripcion + "), RUT " + ((System.String)oGrid.DataTable.GetValue("U_RUT", i)).Trim() + ", Tipo doc " + ((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim() + ", Folio " + ((System.Int32)oGrid.DataTable.GetValue("U_Folio", i)).ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    OutLog("No se actualizado en el portal(" + sDescripcion + "), RUT " + ((System.String)oGrid.DataTable.GetValue("U_RUT", i)).Trim() + ", Tipo doc " + ((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim() + ", Folio " + ((System.Int32)oGrid.DataTable.GetValue("U_Folio", i)).ToString());
                                }
                            }
                        }
                        else
                        {
                            FSBOApp.StatusBar.SetText("No se actualizado en el portal(" + sDescripcion + "), RUT " + ((System.String)oGrid.DataTable.GetValue("U_RUT", i)).Trim() + ", Tipo doc " + ((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim() + ", Folio " + ((System.Int32)oGrid.DataTable.GetValue("U_Folio", i)).ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            OutLog("No se actualizado en el portal(" + sDescripcion + "), RUT " + ((System.String)oGrid.DataTable.GetValue("U_RUT", i)).Trim() + ", Tipo doc " + ((System.String)oGrid.DataTable.GetValue("U_TipoDoc", i)).Trim() + ", Folio " + ((System.Int32)oGrid.DataTable.GetValue("U_Folio", i)).ToString());
                        }
                    }
                }
                return true;
            }
            catch (Exception x)
            {
                FSBOApp.StatusBar.SetText("ActualizarEstado: " + x.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("ActualizarEstado: " + x.Message + " ** Trace: " + x.StackTrace);
                return false;
            }
        }


        private void MostrarPDF(Int32 Linea)
        {
            String Code;
            String sXml;
            String TipoDoc;
            String Folio;
            String RUTEmisor;
            String oPath;
            String sNombreArchivo;
            String sNombrePDF;
            ReportDocument rpt = new ReportDocument();
            ConnectionInfo connection = new ConnectionInfo();
            TableLogOnInfo logOnInfo;
            Boolean flag = true;
            String Pass = "";

            try
            {
                oGrid = (Grid)(oForm.Items.Item("grid").Specific);
                Code = Convert.ToString(((System.Int32)oGrid.DataTable.GetValue("DocEntry", Linea)), _nf);
                sXml = ((System.String)oGrid.DataTable.GetValue("U_Xml", Linea));
                TipoDoc = ((System.String)oGrid.DataTable.GetValue("U_TipoDoc", Linea));
                RUTEmisor = ((System.String)oGrid.DataTable.GetValue("U_RUT", Linea));
                Folio = Convert.ToString(((System.Int32)oGrid.DataTable.GetValue("U_Folio", Linea)), _nf);
                if (sXml != "")
                {
                    oPath = System.IO.Path.GetDirectoryName(TMultiFunctions.ParamStr(0));
                    try
                    {
                        if (GlobalSettings.RunningUnderSQLServer)
                            sNombreArchivo = oPath + "\\Reports\\CL\\SQL\\ReporteXML.rpt";
                        else
                            sNombreArchivo = oPath + "\\Reports\\CL\\HANA\\ReporteXML.rpt";
                        sNombrePDF = oPath + @"\PDF\" + RUTEmisor + "_" + TipoDoc + "_" + Folio + ".pdf";
                        if (File.Exists(sNombrePDF))
                        {
                            System.Diagnostics.Process proc = new System.Diagnostics.Process();
                            proc.StartInfo.FileName = sNombrePDF;
                            proc.Start();
                        }
                        else
                        {

                            FSBOf.AddRepKey(Code, "FEREPORTXML", "FEREPORTXML");//oForm.TypeEx);
                            GlobalSettings.CrystalReportFileName = sNombreArchivo;
                            try
                            {
                                FSBOApp.Menus.Item("4873").Activate();
                            }
                            catch { }

                            /*FSBOApp.Menus.Item("4873").Activate();
                            var oFormB = FSBOApp.Forms.ActiveForm;
                            ((EditText)oFormB.Items.Item("410000004").Specific).Value = sNombreArchivo;
                            oFormB.Items.Item("410000005").Click(BoCellClickType.ct_Regular);*/
                        }
                    }
                    catch (Exception p)
                    {
                        FSBOApp.StatusBar.SetText(p.Message + " ** Trace: " + p.StackTrace, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        OutLog("Cargar Crystal: " + p.Message + " ** Trace: " + p.StackTrace);
                    }
                }
                else
                    FSBOApp.StatusBar.SetText("No se ha encontrado xml que genera PDF", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
            }
            catch (Exception x)
            {
                FSBOApp.StatusBar.SetText(x.Message + " ** Trace: " + x.StackTrace, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("MostrarPDF: " + x.Message + " ** Trace: " + x.StackTrace);
            }
        }

        private void AgregarMensajeGridResumen(string sMensaje, string sDocNum = "", bool bError = false)
        {
            try
            {
                oGrid2 = (Grid)(oForm.Items.Item("grid2").Specific);
                oDTRES.Rows.Add();
                int Fila = oDTRES.Rows.Count - 1;

                oDTRES.SetValue("Descripcion", Fila, sMensaje);
                oDTRES.SetValue("Fecha / Hora", Fila, DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                if (sDocNum.Trim().Length > 0)
                    oDTRES.SetValue("Documento", Fila, sDocNum);
                if (bError)
                    oGrid2.CommonSetting.SetCellBackColor(Fila + 1, 6, ColorTranslator.ToOle(Color.Red));

                SAPbouiCOM.RowHeaders oHeader = null;
                oHeader = oGrid2.RowHeaders;
                //Enumera Fila
                oHeader.SetText(Fila, Convert.ToString(Fila + 1));
            }
            catch { }

        }

        private void CrearDocto(Int32 DocEntry, String RUT, Int32 FolioNum, String TipoDoc, int nFilaGrid)
        {
            SAPbobsCOM.Recordset ors = ((SAPbobsCOM.Recordset)FCmpny.GetBusinessObject(BoObjectTypes.BoRecordset));
            SAPbobsCOM.Recordset orsAux = ((SAPbobsCOM.Recordset)FCmpny.GetBusinessObject(BoObjectTypes.BoRecordset));
            String CardCode;
            Int32 OC;
            Int32 lRetCode;
            Int32 nErr;
            String sErr;
            String DocACrear;
            String CodDocRef;
            SAPbobsCOM.Documents oDocumentsOC;
            SAPbobsCOM.Documents oDocuments;
            bool FactCompraCreado = false;
            string CrearBaseEM = oDTParams.GetValue("EntMer", 0).ToString().Trim();
            string sEMCode = oDTParams.GetValue("CodEM", 0).ToString().Trim();
            string sTipoBase = CrearBaseEM == "Y" ? "Entrada de Mercancia" : "Orden de Compra";

            string DocRefAct = CrearBaseEM == "Y" ? oGrid.DataTable.GetValue("EM", nFilaGrid).ToString() : oGrid.DataTable.GetValue("OC", nFilaGrid).ToString();
            string DocRefOrig = CrearBaseEM == "Y" ? oGrid.DataTable.GetValue("EMOri", nFilaGrid).ToString() : oGrid.DataTable.GetValue("OCOri", nFilaGrid).ToString();
            bool esDocRefXML = DocRefAct != DocRefOrig ? false : true; //Indica si OC / EM fue ingresada por el usuario
            esDocRefXML = false; //No se aplicaran las validacion en el XML del Documento solo solbre el Documento de SAP

            try
            {
                if (GlobalSettings.RunningUnderSQLServer)
                    s = @"SELECT ISNULL(U_CrearDocC,'N') 'Crear', ISNULL(U_FProv,'Y') 'FProv' FROM [@VID_FEPARAM]";
                else
                    s = @"SELECT IFNULL(""U_CrearDocC"",'N') ""Crear"", IFNULL(""U_FProv"",'Y') ""FProv"" FROM ""@VID_FEPARAM"" ";
                ors.DoQuery(s);

                if (((System.String)ors.Fields.Item("Crear").Value).Trim() == "Y")
                {
                    if (((System.String)ors.Fields.Item("FProv").Value).Trim() == "Y")
                        DocACrear = "P"; //Preliminar
                    else
                        DocACrear = "R"; //Real

                    if (GlobalSettings.RunningUnderSQLServer)
                        s = @"SELECT CardCode FROM OCRD WHERE REPLACE(LicTradNum,'.','') = '{0}' AND CardType = 'S' AND frozenFor = 'N'";
                    else
                        s = @"SELECT ""CardCode"" FROM ""OCRD"" WHERE REPLACE(""LicTradNum"",'.','') = '{0}' AND ""CardType"" = 'S' AND ""frozenFor"" = 'N'";
                    s = String.Format(s, RUT.Replace(".", ""));
                    ors.DoQuery(s);
                    if (ors.RecordCount == 0)
                    {
                        FSBOApp.StatusBar.SetText("No se ha encontrado proveedor en el Maestro SN, RUT " + RUT, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                        AgregarMensajeGridResumen("No se ha encontrado proveedor en el Maestro SN, RUT " + RUT);
                    }
                    else
                    {
                        CardCode = ((System.String)ors.Fields.Item("CardCode").Value).Trim();
                        CodDocRef = CrearBaseEM == "Y" ? sEMCode : "801";

                        if (GlobalSettings.RunningUnderSQLServer)
                            s = @"SELECT DISTINCT T0.U_CodRef, T0.U_Folio, convert(nvarchar(max), T1.U_Xml) 'U_XML'
                                      FROM [@VID_FEDTECPRAD] T0
                                      JOIN [@VID_FEDTECPRA] T1 ON T1.DocEntry = T0.DocEntry
                                     WHERE T0.DocEntry = {0}
                                       AND T0.U_CodRef = '" + CodDocRef + @"'
                                       ";
                        else
                            s = @"SELECT DISTINCT T0.""U_CodRef"", T0.""U_Folio"", cast(BINTOSTR(cast(T1.""U_Xml"" as binary)) as varchar) as ""U_Xml""
                                      FROM ""@VID_FEDTECPRAD"" T0
                                      JOIN ""@VID_FEDTECPRA"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                                     WHERE T0.""DocEntry"" = {0}
                                       AND T0.""U_CodRef"" = '" + CodDocRef + "'";
                        s = String.Format(s, DocEntry);
                        ors.DoQuery(s);
                        if (ors.RecordCount == 0 && esDocRefXML)
                        {
                            FSBOApp.StatusBar.SetText("No se ha encontrado " + sTipoBase + " para la factura " + FolioNum.ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                            AgregarMensajeGridResumen("No se ha encontrado " + sTipoBase + " para la factura " + FolioNum.ToString(), "", true);
                        }
                        else
                        {
                            if (ors.RecordCount > 1 && esDocRefXML)
                            {
                                FSBOApp.StatusBar.SetText("Documento posee mas de una " + sTipoBase + " -> " + FolioNum.ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                AgregarMensajeGridResumen("Documento posee mas de una " + sTipoBase + " -> " + FolioNum.ToString(), "", true);
                                return;
                            }

                            if (((System.String)ors.Fields.Item("U_Xml").Value).Trim() == "" && esDocRefXML)
                            {
                                FSBOApp.StatusBar.SetText("Documento no posee XML -> " + FolioNum.ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                AgregarMensajeGridResumen("Documento no posee XML -> " + FolioNum.ToString(), "", true);
                                return;
                            }

                            var sOC = esDocRefXML ? ((System.String)ors.Fields.Item("U_Folio").Value).Trim() : DocRefAct;
                            if (Int32.TryParse(sOC, out OC))
                            {
                                if (CrearBaseEM != "Y") // OC
                                {
                                    if (GlobalSettings.RunningUnderSQLServer)
                                        s = @"SELECT T0.DocEntry, T0.DocStatus, T0.DocTotal, T0.VatSum, COUNT(*) 'Cant'
                                            FROM OPOR T0
                                            JOIN POR1 T1 ON T1.DocEntry = T0.DocEntry
                                            WHERE T0.DocNum = {0}
                                            AND T0.CardCode = '{1}'
                                            GROUP BY T0.DocEntry, T0.DocStatus, T0.DocTotal, T0.VatSum, T0.DocDate";
                                    else
                                        s = @"SELECT T0.""DocEntry"", T0.""DocStatus"", T0.""DocTotal"", T0.""VatSum"", COUNT(*) ""Cant""
                                            FROM ""OPOR"" T0
                                            JOIN ""POR1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                                            WHERE T0.""DocNum"" = {0}
                                            AND T0.""CardCode"" = '{1}'
                                            GROUP BY T0.""DocEntry"", T0.""DocStatus"", T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"" ";
                                }
                                else // EM
                                {
                                    string sBuscarEMDocNum = oDTParams.GetValue("BuscarEMDocNum", 0).ToString().Trim();
                                    if (sBuscarEMDocNum == "Y") //buscar EM por DocNum
                                    {
                                        if (GlobalSettings.RunningUnderSQLServer)
                                            s = @"SELECT T0.DocEntry, T0.DocStatus, T0.DocTotal, T0.VatSum, COUNT(*) 'Cant'
                                            FROM OPDN T0
                                            JOIN PDN1 T1 ON T1.DocEntry = T0.DocEntry
                                            WHERE T0.DocNum = {0}
                                            AND T0.CardCode = '{1}'
                                            GROUP BY T0.DocEntry, T0.DocStatus, T0.DocTotal, T0.VatSum, T0.DocDate";
                                        else
                                            s = @"SELECT T0.""DocEntry"", T0.""DocStatus"", T0.""DocTotal"", T0.""VatSum"", COUNT(*) ""Cant""
                                            FROM ""OPDN"" T0
                                            JOIN ""PDN1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                                            WHERE T0.""DocNum"" = {0}
                                            AND T0.""CardCode"" = '{1}'
                                            GROUP BY T0.""DocEntry"", T0.""DocStatus"", T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"" ";
                                    }
                                    else   //buscar EM por FolioNum
                                    {
                                        if (GlobalSettings.RunningUnderSQLServer)
                                            s = @"SELECT T0.DocEntry, T0.DocStatus, T0.DocTotal, T0.VatSum, COUNT(*) 'Cant'
                                            FROM OPDN T0
                                            JOIN PDN1 T1 ON T1.DocEntry = T0.DocEntry
                                            WHERE T0.FolioNum = {0}
                                            AND T0.CardCode = '{1}'
                                            GROUP BY T0.DocEntry, T0.DocStatus, T0.DocTotal, T0.VatSum, T0.DocDate";
                                        else
                                            s = @"SELECT T0.""DocEntry"", T0.""DocStatus"", T0.""DocTotal"", T0.""VatSum"", COUNT(*) ""Cant""
                                            FROM ""OPDN"" T0
                                            JOIN ""PDN1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                                            WHERE T0.""FolioNum"" = {0}
                                            AND T0.""CardCode"" = '{1}'
                                            GROUP BY T0.""DocEntry"", T0.""DocStatus"", T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"" ";
                                    }
                                }
                                s = String.Format(s, OC, CardCode);
                                ors.DoQuery(s);
                                if (ors.RecordCount == 0)
                                {
                                    FSBOApp.StatusBar.SetText("No se ha encontrado " + sTipoBase + " en SAP -> " + sOC, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    AgregarMensajeGridResumen("No se ha encontrado " + sTipoBase + " en SAP -> " + sOC, "", true);
                                    return;
                                }
                                else
                                {
                                    var CantLineasOC = ((System.Int32)ors.Fields.Item("Cant").Value);
                                    var OCDocEntry = ((System.Int32)ors.Fields.Item("DocEntry").Value);
                                    var EMDocEntry = 0;
                                    var EMDocNum = 0;
                                    var OCDocStatus = ((System.String)ors.Fields.Item("DocStatus").Value).Trim();
                                    var OCDocTotal = ((System.Double)ors.Fields.Item("DocTotal").Value);
                                    var OCVatSum = ((System.Double)ors.Fields.Item("VatSum").Value);
                                    var BaseType = 22;
                                    //Detalle
                                    if (GlobalSettings.RunningUnderSQLServer)
                                        s = @"SELECT COUNT(*) 'Cant'
                                                                    FROM [@VID_FEXMLCD]
                                                                    WHERE Code = '{0}'";
                                    else
                                        s = @"SELECT COUNT(*) ""Cant""
                                                                    FROM ""@VID_FEXMLCD""
                                                                    WHERE ""Code"" = '{0}'";
                                    s = String.Format(s, DocEntry);
                                    ors.DoQuery(s);
                                    var CantLinFE = ((System.Int32)ors.Fields.Item("Cant").Value);

                                    //if (CantLineasOC != CantLinFE)
                                    //{
                                    //    FSBOApp.StatusBar.SetText("Cantidad de lineas entre " + sTipoBase + " y Documento Elec. son diferentes", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    //    return;
                                    //}

                                    //mejorar query para que muestre toda la OC con sus entregas para dejar todo cerrado con la factura***************************
                                    //cabecera

                                    if (GlobalSettings.RunningUnderSQLServer)
                                        s = @"SELECT U_FchEmis, U_FchVenc, U_RznSoc, U_MntNeto, U_MntExe, U_MntTotal, U_IVA
                                                      FROM [@VID_FEXMLC] 
                                                     WHERE Code = '{0}'";
                                    else
                                        s = @"SELECT ""U_FchEmis"", ""U_FchVenc"", ""U_RznSoc"", ""U_MntNeto"", ""U_MntExe"", ""U_MntTotal"", ""U_IVA""
                                                      FROM ""@VID_FEXMLC""
                                                     WHERE ""Code"" = '{0}'";
                                    s = String.Format(s, DocEntry);
                                    orsAux.DoQuery(s);
                                    var FchEmis = ((System.DateTime)orsAux.Fields.Item("U_FchEmis").Value);
                                    var Fchvenc = ((System.DateTime)orsAux.Fields.Item("U_FchVenc").Value);
                                    var RznSoc = ((System.String)orsAux.Fields.Item("U_RznSoc").Value).Trim();
                                    var MntNeto = ((System.Double)orsAux.Fields.Item("U_MntNeto").Value);
                                    var MntExe = ((System.Double)orsAux.Fields.Item("U_MntExe").Value);
                                    var MntTotal = ((System.Double)orsAux.Fields.Item("U_MntTotal").Value);
                                    var IVA = ((System.Double)orsAux.Fields.Item("U_IVA").Value);

                                    if (CrearBaseEM != "Y") // OC
                                    {
                                        //verificar si solo existe oc, la fe se crea a partir de oc, si tiene entreda de mercancia se crea a partir de la entrada de mercancia 
                                        if (GlobalSettings.RunningUnderSQLServer)
                                            s = @"SELECT T0.DocEntry, T1.LineNum, T1.ObjType, T1.OpenQty 'Quantity'
                                              FROM OPOR T0
                                              JOIN POR1 T1 ON T1.DocEntry = T0.DocEntry
                                             WHERE 1=1
                                               AND T0.DocNum = {0}
                                            UNION 
                                            SELECT P1.DocEntry, P1.LineNum, P1.ObjType, P1.Quantity--, P1.BaseLine
                                              FROM OPOR T0
                                              JOIN POR1 T1 ON T1.DocEntry = T0.DocEntry
                                              JOIN PDN1 P1 ON P1.BaseEntry = T1.DocEntry
                                                          AND P1.BaseType = T0.ObjType
			                                              AND P1.BaseLine = T1.LineNum
                                             WHERE 1=1
                                               AND T0.DocNum = {0}";
                                        else
                                            s = @"SELECT T0.""DocEntry"", T1.""LineNum"", T1.""ObjType"", T1.""OpenQty"" ""Quantity""
                                              FROM ""OPOR"" T0
                                              JOIN ""POR1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                                             WHERE 1=1
                                               AND T0.""DocNum"" = {0}
                                            UNION 
                                            SELECT P1.""DocEntry"", P1.""LineNum"", P1.""ObjType"", P1.""Quantity"" --, P1.BaseLine
                                              FROM ""OPOR"" T0
                                              JOIN ""POR1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                                              JOIN ""PDN1"" P1 ON P1.""BaseEntry"" = T1.""DocEntry""
                                                          AND P1.""BaseType"" = T0.""ObjType""
			                                              AND P1.""BaseLine"" = T1.""LineNum""
                                             WHERE 1=1
                                               AND T0.""DocNum"" = {0}";
                                    }
                                    else //EM
                                    {
                                        string sBuscarEMDocNum = oDTParams.GetValue("BuscarEMDocNum", 0).ToString().Trim();
                                        if (sBuscarEMDocNum == "Y") //buscar EM por DocNum
                                        {
                                            if (GlobalSettings.RunningUnderSQLServer)
                                                s = @"SELECT T0.DocEntry, T1.LineNum, T1.ObjType, T1.OpenQty 'Quantity'
                                              FROM OPDN T0
                                              JOIN PDN1 T1 ON T1.DocEntry = T0.DocEntry
                                             WHERE 1=1
                                               AND T0.DocNum = {0}";
                                            else
                                                s = @"SELECT T0.""DocEntry"", T1.""LineNum"", T1.""ObjType"", T1.""OpenQty"" ""Quantity""
                                              FROM ""OPDN"" T0
                                              JOIN ""PDN1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                                             WHERE 1=1
                                               AND T0.""DocNum"" = {0}";
                                        }
                                        else   //Buscar EM por FolioNum
                                        {
                                            if (GlobalSettings.RunningUnderSQLServer)
                                                s = @"SELECT T0.DocEntry, T1.LineNum, T1.ObjType, T1.OpenQty 'Quantity'
                                              FROM OPDN T0
                                              JOIN PDN1 T1 ON T1.DocEntry = T0.DocEntry
                                             WHERE 1=1
                                               AND T0.FolioNum = {0}";
                                            else
                                                s = @"SELECT T0.""DocEntry"", T1.""LineNum"", T1.""ObjType"", T1.""OpenQty"" ""Quantity""
                                              FROM ""OPDN"" T0
                                              JOIN ""PDN1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                                             WHERE 1=1
                                               AND T0.""FolioNum"" = {0}";
                                        }
                                    }
                                    s = String.Format(s, OC);
                                    orsAux.DoQuery(s);

                                    if (orsAux.RecordCount > 0)
                                    {
                                        //como no salio por algun return sigo con la creacion del documento en SAP
                                        var men = "";
                                        if (DocACrear == "P") //preliminar dentro de lac configuracion
                                        {
                                            oDocuments = ((SAPbobsCOM.Documents)FCmpny.GetBusinessObject(BoObjectTypes.oDrafts));
                                            men = "Factura Preliminar"; //preliminar dentro de lac configuracion
                                        }
                                        else
                                        {
                                            oDocuments = ((SAPbobsCOM.Documents)FCmpny.GetBusinessObject(BoObjectTypes.oPurchaseInvoices));
                                            men = "Factura";
                                        }

                                        oDocuments.CardCode = CardCode;
                                        oDocuments.CardName = RznSoc;
                                        //oDocuments.DocDate = FchEmis;
                                        var date = DateTime.Now; 
                                        oDocuments.DocDate = DateTime.Now;
                                        oDocuments.DocDueDate = Fchvenc;
                                        oDocuments.FolioPrefixString = TipoDoc;
                                        oDocuments.FolioNumber = FolioNum;
                                        oDocuments.Comments = "Creado por addon FE en Aceptación del DTE,  Basado en " + sTipoBase + " " + OC + ".";
                                        if (DocACrear == "P")
                                            oDocuments.DocObjectCode = BoObjectTypes.oPurchaseInvoices;

                                        while (!orsAux.EoF)
                                        {
                                            oDocuments.Lines.BaseEntry = ((System.Int32)orsAux.Fields.Item("DocEntry").Value);
                                            oDocuments.Lines.BaseLine = ((System.Int32)orsAux.Fields.Item("LineNum").Value);
                                            oDocuments.Lines.BaseType = Convert.ToInt32(((System.String)orsAux.Fields.Item("ObjType").Value), _nf);
                                            oDocuments.Lines.Quantity = ((System.Double)orsAux.Fields.Item("Quantity").Value);
                                            oDocuments.Lines.Add();
                                            orsAux.MoveNext();
                                        }

                                        //oDocuments.VatSum = 0;
                                        oDocuments.DocTotal = MntTotal;

                                        lRetCode = oDocuments.Add();
                                        if (lRetCode != 0)
                                        {
                                            FCmpny.GetLastError(out nErr, out sErr);
                                            OutLog("No se ha creado documento en SAP, " + men + " -> " + FolioNum.ToString() + " - " + sErr);
                                            FSBOApp.StatusBar.SetText("No se ha creado documento en SAP, " + men + " -> " + FolioNum.ToString() + " - " + sErr, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                            AgregarMensajeGridResumen("No se ha creado documento en SAP, " + men + " -> " + FolioNum.ToString() + " - " + sErr, "", true);
                                        }
                                        else
                                        {
                                            FSBOApp.StatusBar.SetText("Se ha creado satisfactoriamente el documento en SAP, " + men + " -> " + FolioNum.ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                            FactCompraCreado = true;
                                            //guardar registro para monitor de registros creados
                                            var NuevoDocEntry = FCmpny.GetNewObjectKey();
                                            if (GlobalSettings.RunningUnderSQLServer)
                                                s = @"UPDATE [@VID_FEDTECPRA] SET U_DocEntry = {1}, U_ObjType = '{2}' WHERE DocEntry = {0}";
                                            else
                                                s = @"UPDATE ""@VID_FEDTECPRA"" SET ""U_DocEntry"" = {1}, ""U_ObjType"" = '{2}' WHERE ""DocEntry"" = {0}";
                                            s = String.Format(s, DocEntry, NuevoDocEntry, (DocACrear == "P" ? "112" : "18"));
                                            orsAux.DoQuery(s);
                                            AgregarMensajeGridResumen("Se ha creado satisfactoriamente el documento en SAP, " + men + " -> " + FolioNum.ToString(), NuevoDocEntry);

                                        }
                                    }
                                    else
                                    {
                                        FSBOApp.StatusBar.SetText("No se ha encontrado detalle " + sTipoBase + " " + OC, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        AgregarMensajeGridResumen("No se ha encontrado detalle " + sTipoBase + " " + OC, "", true);
                                        return;
                                    }
                                }
                            }
                            else
                            {
                                FSBOApp.StatusBar.SetText("Numero " + sTipoBase + " no es valido -> " + ((System.String)ors.Fields.Item("U_Folio").Value), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                AgregarMensajeGridResumen("Numero " + sTipoBase + " no es valido -> " + ((System.String)ors.Fields.Item("U_Folio").Value), "", true);
                                return;
                            }

                        }
                        //Varifica, si no se crea la factura de compra
                        if (FactCompraCreado == false)
                        {
                            if (GlobalSettings.RunningUnderSQLServer)
                                s = @"UPDATE [@VID_FEDTECPRA] SET U_Validacion = '<FC NO CREADA>' WHERE DocEntry = {0}";
                            else
                                s = @"UPDATE ""@VID_FEDTECPRA"" SET ""U_Validacion"" = '<FC NO CREADA>' WHERE ""DocEntry"" = {0}";
                            s = String.Format(s, DocEntry.ToString());
                            orsAux.DoQuery(s);
                        }
                    }
                }

            }
            catch (Exception x)
            {
                FSBOApp.StatusBar.SetText("CrearDocto: " + x.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("CrearDocto: " + x.Message + " ** Trace: " + x.StackTrace);
            }
            finally
            {
                FSBOf._ReleaseCOMObject(ors);
                FSBOf._ReleaseCOMObject(orsAux);
            }
        }

        private bool ValidaDocRefProveedor(string RutProv, string CodRef, string NumRef)
        {
            bool Existe = false;
            try
            {
                SAPbobsCOM.Recordset ors = ((SAPbobsCOM.Recordset)FCmpny.GetBusinessObject(BoObjectTypes.BoRecordset));

                if (GlobalSettings.RunningUnderSQLServer)
                    s = @"SELECT  U_RUT, U_Folio FROM [@VID_FEDTECPRA]  WHERE U_RUT = '{0}' AND U_CodRefGen  = '{1}' AND U_FolioRefGen  = '{2}' ";
                else
                    s = @"SELECT  ""U_RUT"", ""U_Folio"" FROM ""@VID_FEDTECPRA"" WHERE ""U_RUT"" = '{0}' AND ""U_CodRefGen""  = '{1}' AND ""U_FolioRefGen""  = '{2}'";
                s = String.Format(s, RutProv, CodRef, NumRef);
                ors.DoQuery(s);

                if (ors.RecordCount > 0)
                    Existe = true;

            }
            catch { }

            return Existe;

        }

        private string ValidaDocEstado(string CodRef, string NumRef)
        {
            string respuesta = "";
            try
            {
                SAPbobsCOM.Recordset orsaux = ((SAPbobsCOM.Recordset)FCmpny.GetBusinessObject(BoObjectTypes.BoRecordset));
                string TablaC = CodRef == "OC" ? "OPOR" : "OPDN" ;
                string TablaD = CodRef == "OC" ? "POR1" : "PDN1";
                string sBuscarEMDocNum = oDTParams.GetValue("BuscarEMDocNum", 0).ToString().Trim();
                if (sBuscarEMDocNum == "Y") //buscar EM por DocNum
                {
                    //Busca datos de la OC/EM
                    if (GlobalSettings.RunningUnderSQLServer)
                        s = @"SELECT T0.DocEntry, T0.CardCode, T0.DocStatus, T0.Confirmed, T0.CANCELED, T0.DocTotal, T0.VatSum, T0.DocDate, COUNT(*) 'Cant'
                                                        FROM {0} T0
                                                        JOIN {1} T1 ON T1.DocEntry = T0.DocEntry
                                                        WHERE CAST(T0.DocNum as NVARCHAR(30)) = '{2}'
                                                        GROUP BY T0.DocEntry, T0.CardCode, T0.DocStatus,  T0.CANCELED, T0.Confirmed, T0.DocTotal, T0.VatSum, T0.DocDate";
                    else
                        s = @"SELECT T0.""DocEntry"", T0.""CardCode"", T0.""DocStatus"", T0.""Confirmed"", T0.""CANCELED"",  T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"", COUNT(*) ""Cant""
                                                        FROM ""{0}"" T0
                                                        JOIN ""{1}"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                                                        WHERE CAST(T0.""DocNum"" as NVARCHAR(30)) = '{2}'
                                                        GROUP BY T0.""DocEntry"", T0.""CardCode"", T0.""DocStatus"",  T0.""CANCELED"",  T0.""Confirmed"",T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"" ";
                    s = String.Format(s, TablaC, TablaD, NumRef);
                }
                else //buscar EM por FolioNum
                {
                    if (GlobalSettings.RunningUnderSQLServer)
                        s = @"SELECT T0.DocEntry, T0.CardCode, T0.DocStatus, T0.Confirmed, T0.CANCELED, T0.DocTotal, T0.VatSum, T0.DocDate, COUNT(*) 'Cant'
                                                        FROM {0} T0
                                                        JOIN {1} T1 ON T1.DocEntry = T0.DocEntry
                                                        WHERE CAST(T0.FolioNum as NVARCHAR(30)) = '{2}'
                                                        GROUP BY T0.DocEntry, T0.CardCode, T0.DocStatus,  T0.CANCELED, T0.Confirmed, T0.DocTotal, T0.VatSum, T0.DocDate";
                    else
                        s = @"SELECT T0.""DocEntry"", T0.""CardCode"", T0.""DocStatus"", T0.""Confirmed"", T0.""CANCELED"",  T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"", COUNT(*) ""Cant""
                                                        FROM ""{0}"" T0
                                                        JOIN ""{1}"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                                                        WHERE CAST(T0.""FolioNum"" as NVARCHAR(30)) = '{2}'
                                                        GROUP BY T0.""DocEntry"", T0.""CardCode"", T0.""DocStatus"",  T0.""CANCELED"",  T0.""Confirmed"",T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"" ";
                    s = String.Format(s, TablaC, TablaD, NumRef);
                }
                orsaux.DoQuery(s);

                if (orsaux.RecordCount == 0)
                {
                    respuesta = "No se ha encontrado Orden de Compra en SAP, ";
                }
                else
                {

                    var OCDocEntry = ((System.Int32)orsaux.Fields.Item("DocEntry").Value);
                    var OCDocStatus = ((System.String)orsaux.Fields.Item("DocStatus").Value).Trim();
                    var OCDocTotal = ((System.Double)orsaux.Fields.Item("DocTotal").Value);
                    var OCVatSum = ((System.Double)orsaux.Fields.Item("VatSum").Value);
                    var OCDocDate = ((System.DateTime)orsaux.Fields.Item("DocDate").Value);
                    var CantLineasOC = ((System.Int32)orsaux.Fields.Item("Cant").Value);
                    var AnuladoOC = ((System.String)orsaux.Fields.Item("CANCELED").Value).Trim();
                    var Autorizada = ((System.String)orsaux.Fields.Item("Confirmed").Value).Trim();

                    if (OCDocStatus != "O")
                    {
                        respuesta = "La " + CodRef + " en SAP esta Cerrada, ";
                    }
                    if (OCDocStatus == "O" && Autorizada == "N")
                    {
                        respuesta = "La " + CodRef + " en SAP No esta Autorizada, ";
                    }
                    if (AnuladoOC == "Y")
                    {
                        respuesta = "La " + CodRef + " en SAP esta Anulada, ";
                    }
                }
            }
            catch { }

            return respuesta;

        }

        private void ValidarDocumentos(bool MostrarMensajes = true)
        {
            SAPbouiCOM.ProgressBar oProgressBar = null;
            SAPbouiCOM.RowHeaders oHeader = null;
            oHeader = oGrid.RowHeaders;

            DateTime dFecRecp = new DateTime();
            Int32 valorFila;
            Int32 nDias;
            string sDias;
            bool bEditable;

            string ListaBlanca = "";
            string ListaNegra = "";
            //Parametros
            string ValidarEM = oDTParams.GetValue("EntMer", 0).ToString().Trim();
            string VisualizarLN = oDTParams.GetValue("VListNegra", 0).ToString().Trim();
            string ModificarLB = oDTParams.GetValue("MListBlanca", 0).ToString().Trim();
            string ModificarLN = oDTParams.GetValue("MListNegra", 0).ToString().Trim();
            string ModificarFCT = oDTParams.GetValue("MFacDifer", 0).ToString().Trim();

            FSBOApp.StatusBar.SetText("Validando Documentos, Por Favor Espere ... ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            oProgressBar = FSBOApp.StatusBar.CreateProgressBar("Validando Documentos, Por Favor Espere ...", oGrid.Rows.Count, false);

            try
            {
                TotDocu = oGrid.Rows.Count;
                TotAcep = 0;
                TotRecl = 0;
                TotSele = 0;

                for (Int32 numfila = 0; numfila <= oGrid.Rows.Count - 1; numfila++)
                {
                    valorFila = oGrid.GetDataTableRowIndex(numfila);
                    oProgressBar.Value += 1;
                    string tipoDoc = oGrid.DataTable.GetValue("U_TipoDoc", valorFila).ToString();

                    if (valorFila != -1)
                    {
                        //if (MostrarMensajes)
                        //    FSBOApp.StatusBar.SetText("Validando Documentos, Por Favor Espere ... " + (numfila + 1).ToString() + "/" + oGrid.Rows.Count.ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                        //VErifica Validacion Antigua (U_VAlidacion) que se modifica en el Servicio
                        //if (((System.String)oGrid.DataTable.GetValue("U_Validacion", valorFila)).Trim() == "OK")
                        //{
                        //    oGrid.CommonSetting.SetCellBackColor(numfila + 1, 6, ColorTranslator.ToOle(Color.LightGreen));
                        //    oGrid.DataTable.SetValue("Acepta", numfila, "Y");
                        //    oGrid.DataTable.SetValue("U_EstadoLey", numfila, "ACD");
                        //}
                        //else
                        //{
                        //    oGrid.CommonSetting.SetCellBackColor(numfila + 1, 6, ColorTranslator.ToOle(Color.FromArgb(240, 128, 128)));
                        //    oGrid.DataTable.SetValue("Acepta", numfila, "N");
                        //    oGrid.DataTable.SetValue("U_EstadoLey", numfila, "RCD");
                        //}

                        //Tiempo Restante
                        try
                        {
                            dFecRecp = DateTime.Parse(oDataTable.GetValue("U_FechaRecep", numfila).ToString()).AddDays(8);
                            // sFecRecp = oDataTable.GetValue("U_FechaRecep", i).ToString();

                            var Trestante = (dFecRecp - DateTime.Now).ToString(@"dd\d\ hh\h\ mm\m\ ").Replace("d", " Dia(s)").Replace("h ", ":").Replace("m", "");
                            var Dias = Convert.ToInt32((dFecRecp - DateTime.Now).TotalDays);
                            //var horas2 = (DateTime.Now - DateTime.Parse(sFecRecp)).ToString(@"dd\d\ hh\h\ mm\m\ ");
                            if (Trestante.Substring(0, 1) == "0")
                                Trestante = Trestante.Remove(0, 1);

                            oDataTable.SetValue("Tiempo Rest.", numfila, Dias < 0 ? "0 Dia(s) 00:00" : Trestante);
                            oDataTable.SetValue("DiasRest", numfila, Dias.ToString());
                        }
                        catch { }

                        // Lista Blanca
                        ListaBlanca = oGrid.DataTable.GetValue("LB", valorFila).ToString().Trim();
                        if (ListaBlanca.Length > 0)
                        {
                            oGrid.CommonSetting.SetCellBackColor(numfila + 1, 6, ColorTranslator.ToOle(Color.LightGreen));
                            oGrid.DataTable.SetValue("Acepta", numfila, "Y");
                            oGrid.DataTable.SetValue("U_EstadoLey", numfila, "ACD");
                            oGrid.DataTable.SetValue("U_Validacion", numfila, "OK");
                            oGrid.CommonSetting.SetCellBackColor(numfila + 1, 6, ColorTranslator.ToOle(Color.LightGreen));
                            oGrid.DataTable.SetValue("Desc_Valida", numfila, "Lista Blanca, ");
                            bEditable = ModificarLB == "Y" ? true : false;
                            oGrid.CommonSetting.SetCellEditable(valorFila + 1, 14, bEditable);
                            TotAcep += 1;
                            TotSele += 1;
                        }

                        // Lista Negra
                        ListaNegra = oGrid.DataTable.GetValue("LN", valorFila).ToString().Trim();
                        if (ListaNegra.Length > 0)
                        {
                            oGrid.CommonSetting.SetCellBackColor(numfila + 1, 6, ColorTranslator.ToOle(Color.FromArgb(240, 128, 128)));
                            oGrid.DataTable.SetValue("Acepta", numfila, "N");
                            oGrid.DataTable.SetValue("U_EstadoLey", numfila, "RCD");
                            oGrid.DataTable.SetValue("U_Validacion", numfila, "LN");
                            oGrid.CommonSetting.SetCellBackColor(numfila + 1, 6, ColorTranslator.ToOle(Color.FromArgb(240, 128, 128)));
                            oGrid.DataTable.SetValue("Desc_Valida", numfila, "Lista Negra, ");
                            bEditable = ModificarLN == "Y" ? true : false;
                            oGrid.CommonSetting.SetCellEditable(valorFila + 1, 14, bEditable);
                            TotRecl += 1;
                        }


                        //colorea en verde para mostrar que tiene PDF
                        if (((System.String)oGrid.DataTable.GetValue("U_Xml", valorFila)).Trim() == "")
                            oGrid.CommonSetting.SetCellFontColor(numfila + 1, 5, ColorTranslator.ToOle(Color.DarkOrange));
                        //oGrid.CommonSetting.SetCellFontStyle(numfila + 1, 5, BoFontStyle.fs_Bold);


                        //Campo OCOri
                        var sOC = oGrid.DataTable.GetValue("OCOri", valorFila).ToString().Trim();
                        var sOCOrig = sOC;
                        double dTotal = (double)oGrid.DataTable.GetValue("U_Monto", valorFila);
                        DateTime FecEmis = (DateTime)oGrid.DataTable.GetValue("U_FechaEmi", valorFila);
                        string CardCode = oGrid.DataTable.GetValue("CardCode", valorFila).ToString();
                        string DocEntry = oGrid.DataTable.GetValue("DocEntry", valorFila).ToString();

                        string sDescValida;
                        string sValidaRes = oGrid.DataTable.GetValue("Desc_Valida", numfila).ToString().Trim().Length > 0 ? oGrid.DataTable.GetValue("Desc_Valida", numfila).ToString() : "";
                        string sValidaOC;

                        if (sOC.Length == 0 && ((System.String)oGrid.DataTable.GetValue("OC", valorFila)).Trim() == "" && ValidarEM != "Y") // NO OCOrig + NO OC + NO LB + NO LN && ListaBlanca.Length == 0 && ListaNegra.Length == 0
                        {
                            oGrid.DataTable.SetValue("Acepta", numfila, "N");
                            //oGrid.DataTable.SetValue("U_EstadoLey", numfila, "RCD");
                            oGrid.DataTable.SetValue("U_Validacion", numfila, "NO OC");
                            oGrid.DataTable.SetValue("Desc_Valida", numfila, sValidaRes + "Sin asignar Orden de Compra SAP, ");
                            oGrid.CommonSetting.SetCellBackColor(numfila + 1, 6, ColorTranslator.ToOle(Color.FromArgb(240, 128, 128)));
                            TotRecl += 1;
                        }

                        if (((System.String)oGrid.DataTable.GetValue("OC", valorFila)).Trim() == "")
                        {
                            oGrid.DataTable.SetValue("OC", numfila, sOC);
                            if (sOC.Length > 0)
                            {
                                sDescValida = oGrid.DataTable.GetValue("Desc_Valida", numfila).ToString();

                                if (ValidarEM != "Y" && ListaBlanca.Length == 0 && ListaNegra.Length == 0) // Valida OC y NO Lista Blanca / Negra 
                                {
                                    sValidaOC = ValidarOC_Xml(DocEntry, CardCode, dTotal, FecEmis, out sValidaRes);
                                    oGrid.DataTable.SetValue("Desc_Valida", numfila, sValidaOC);

                                    if (sValidaOC.Length == 0) //Si viene vacio esta OK
                                    {
                                        oGrid.DataTable.SetValue("Acepta", numfila, "Y");
                                        oGrid.DataTable.SetValue("U_EstadoLey", numfila, "ACD");
                                        oGrid.DataTable.SetValue("U_Validacion", numfila, "OK");
                                        oGrid.CommonSetting.SetCellBackColor(numfila + 1, 6, ColorTranslator.ToOle(Color.LightGreen));
                                        TotAcep += 1;
                                        TotSele += 1;
                                    }
                                    else
                                    {
                                        oGrid.DataTable.SetValue("Acepta", numfila, "N");
                                        oGrid.DataTable.SetValue("U_EstadoLey", numfila, "RCD");
                                        oGrid.DataTable.SetValue("U_Validacion", numfila, sValidaRes);
                                        oGrid.CommonSetting.SetCellBackColor(numfila + 1, 6, ColorTranslator.ToOle(Color.FromArgb(240, 128, 128)));
                                        bEditable = ModificarFCT == "Y" ? true : false;
                                        oGrid.CommonSetting.SetCellEditable(valorFila + 1, 14, bEditable);
                                        TotRecl += 1;
                                    }
                                }
                            }
                        }
                        else
                        {
                            sOC = oGrid.DataTable.GetValue("OC", valorFila).ToString().Trim();
                            if (sOC.Length > 0)
                            {
                                sDescValida = oGrid.DataTable.GetValue("Desc_Valida", numfila).ToString();

                                if (ValidarEM != "Y" && ListaBlanca.Length == 0 && ListaNegra.Length == 0) // Valida OC y NO Lista Blanca / Negra 
                                {
                                    sValidaOC = ValidarOC_Manual(sOC, CardCode, dTotal, FecEmis, out sValidaRes);
                                    oGrid.DataTable.SetValue("Desc_Valida", numfila, sValidaOC);

                                    if (sValidaOC.Length == 0) //Si viene vacio esta OK
                                    {
                                        oGrid.DataTable.SetValue("Acepta", numfila, "Y");
                                        oGrid.DataTable.SetValue("U_EstadoLey", numfila, "ACD");
                                        oGrid.DataTable.SetValue("U_Validacion", numfila, "OK");
                                        oGrid.CommonSetting.SetCellBackColor(numfila + 1, 6, ColorTranslator.ToOle(Color.LightGreen));
                                        TotAcep += 1;
                                        TotSele += 1;
                                    }
                                    else
                                    {
                                        oGrid.DataTable.SetValue("Acepta", numfila, "N");
                                        oGrid.DataTable.SetValue("U_EstadoLey", numfila, "RCD");
                                        oGrid.DataTable.SetValue("U_Validacion", numfila, sValidaRes);
                                        oGrid.CommonSetting.SetCellBackColor(numfila + 1, 6, ColorTranslator.ToOle(Color.FromArgb(240, 128, 128)));
                                        bEditable = ModificarFCT == "Y" ? true : false;
                                        oGrid.CommonSetting.SetCellEditable(valorFila + 1, 14, bEditable);
                                        TotRecl += 1;
                                    }
                                }
                            }
                        }



                        //Campo EMOri
                        var sEM = oGrid.DataTable.GetValue("EMOri", valorFila).ToString().Trim();
                        var sEMOrig = sEM;
                        string sValidaEM;

                        if (sEM.Length == 0 && ((System.String)oGrid.DataTable.GetValue("EM", valorFila)).Trim() == "" && ValidarEM == "Y") // NO OCOrig + NO OC + NO LB + NO LN  && ListaBlanca.Length == 0 && ListaNegra.Length == 0
                        {
                            oGrid.DataTable.SetValue("Acepta", numfila, "N");
                            //oGrid.DataTable.SetValue("U_EstadoLey", numfila, "RCD");
                            oGrid.DataTable.SetValue("U_Validacion", numfila, "NO EM");
                            oGrid.DataTable.SetValue("Desc_Valida", numfila, sValidaRes + "Sin asignar Entrada Mercancia SAP, ");
                            oGrid.CommonSetting.SetCellBackColor(numfila + 1, 6, ColorTranslator.ToOle(Color.FromArgb(240, 128, 128)));
                            TotRecl += 1;
                        }

                        if (((System.String)oGrid.DataTable.GetValue("EM", valorFila)).Trim() == "")
                        {
                            oGrid.DataTable.SetValue("EM", numfila, sEM);
                            if (sEM.Length > 0)
                            {
                                sDescValida = oGrid.DataTable.GetValue("Desc_Valida", numfila).ToString();

                                if (ValidarEM == "Y" && ListaBlanca.Length == 0 && ListaNegra.Length == 0) // Valida EM y NO Lista Blanca / Negra
                                {
                                    sValidaEM = ValidarEM_Xml(DocEntry, CardCode, dTotal, FecEmis, out sValidaRes);
                                    oGrid.DataTable.SetValue("Desc_Valida", numfila, sValidaEM);

                                    if (sValidaEM.Length == 0) //Si viene vacio esta OK
                                    {
                                        oGrid.DataTable.SetValue("Acepta", numfila, "Y");
                                        oGrid.DataTable.SetValue("U_EstadoLey", numfila, "ACD");
                                        oGrid.DataTable.SetValue("U_Validacion", numfila, "OK");
                                        oGrid.CommonSetting.SetCellBackColor(numfila + 1, 6, ColorTranslator.ToOle(Color.LightGreen));
                                        TotAcep += 1;
                                        TotSele += 1;
                                    }
                                    else
                                    {
                                        oGrid.DataTable.SetValue("Acepta", numfila, "N");
                                        oGrid.DataTable.SetValue("U_EstadoLey", numfila, "RCD");
                                        oGrid.DataTable.SetValue("U_Validacion", numfila, sValidaRes);
                                        oGrid.CommonSetting.SetCellBackColor(numfila + 1, 6, ColorTranslator.ToOle(Color.FromArgb(240, 128, 128)));
                                        bEditable = ModificarFCT == "Y" ? true : false;
                                        oGrid.CommonSetting.SetCellEditable(valorFila + 1, 14, bEditable);
                                        TotRecl += 1;

                                    }
                                }
                            }
                        }
                        else
                        {
                            sEM = oGrid.DataTable.GetValue("EM", valorFila).ToString().Trim();
                            if (sEM.Length > 0)
                            {
                                sDescValida = oGrid.DataTable.GetValue("Desc_Valida", numfila).ToString();

                                if (ValidarEM == "Y" && ListaBlanca.Length == 0 && ListaNegra.Length == 0) // Valida EM y NO Lista Blanca / Negra
                                {
                                    sValidaEM = ValidarEM_Manual(sEM, CardCode, dTotal, FecEmis, out sValidaRes);
                                    oGrid.DataTable.SetValue("Desc_Valida", numfila, sValidaEM);

                                    if (sValidaEM.Length == 0) //Si viene vacio esta OK
                                    {
                                        oGrid.DataTable.SetValue("Acepta", numfila, "Y");
                                        oGrid.DataTable.SetValue("U_EstadoLey", numfila, "ACD");
                                        oGrid.DataTable.SetValue("U_Validacion", numfila, "OK");
                                        oGrid.CommonSetting.SetCellBackColor(numfila + 1, 6, ColorTranslator.ToOle(Color.LightGreen));
                                        TotAcep += 1;
                                        TotSele += 1;
                                    }
                                    else
                                    {
                                        oGrid.DataTable.SetValue("Acepta", numfila, "N");
                                        oGrid.DataTable.SetValue("U_EstadoLey", numfila, "RCD");
                                        oGrid.DataTable.SetValue("U_Validacion", numfila, sValidaRes);
                                        oGrid.CommonSetting.SetCellBackColor(numfila + 1, 6, ColorTranslator.ToOle(Color.FromArgb(240, 128, 128)));
                                        bEditable = ModificarFCT == "Y" ? true : false;
                                        oGrid.CommonSetting.SetCellEditable(valorFila + 1, 14, bEditable);
                                        TotRecl += 1;
                                    }
                                }
                            }
                        }

                        //Valida Forma de Pagos 1 = CONTADO / 3 = ENTREGA GRATUITA
                        string sFormPag = oGrid.DataTable.GetValue("U_FmaPago", numfila).ToString();
                        if (sFormPag != "2" && (tipoDoc == "33" || tipoDoc == "34"))
                        {
                            if (sFormPag == "1" || sFormPag == "3")
                            {
                                //string sDescv = oGrid.DataTable.GetValue("Desc_Valida", numfila).ToString().Trim();
                                //sDescv = sDescv.Trim().Length > 0 ? sDescv + " " : sDescv;
                                //string sValid = sDescv + (sFormPag == "1" ? "Forma de Pago Contado, Integrar a Sistema, " : "Forma de Pago Entrega Gratuita, Integrar a Sistema, ");
                                //oGrid.DataTable.SetValue("Acepta", numfila, "Y");
                                oGrid.DataTable.SetValue("U_EstadoLey", numfila, "");
                                //oGrid.DataTable.SetValue("U_Validacion", numfila, "CONTADO");
                                // oGrid.DataTable.SetValue("Desc_Valida", numfila, sValid);
                                //oGrid.CommonSetting.SetCellBackColor(numfila + 1, 6, ColorTranslator.ToOle(Color.LightGreen));
                                oGrid.CommonSetting.SetCellEditable(valorFila + 1, 14, false);
                            }
                        }
                        else
                        {
                            if (tipoDoc == "61" || tipoDoc == "56")
                            {
                                oGrid.DataTable.SetValue("U_EstadoLey", numfila, "");
                                oGrid.CommonSetting.SetCellEditable(valorFila + 1, 14, false);
                            }
                                
                        }

                        //Colorea en semaforo para dias restantes
                        sDias = oGrid.DataTable.GetValue("DiasRest", valorFila).ToString().Trim();
                        nDias = sDias.Length == 0 ? 0 : Int32.Parse(sDias);

                        if (nDias >= 6)
                            oGrid.CommonSetting.SetCellFontColor(numfila + 1, 8, ColorTranslator.ToOle(Color.Green));
                        else if (nDias == 5 || nDias == 4)
                            oGrid.CommonSetting.SetCellFontColor(numfila + 1, 8, ColorTranslator.ToOle(Color.DarkOrange));
                        else if (nDias < 4)
                            oGrid.CommonSetting.SetCellFontColor(numfila + 1, 8, ColorTranslator.ToOle(Color.Red));

                        //Corrige Fin Descripcion Validacion
                        string sDesc = oGrid.DataTable.GetValue("Desc_Valida", numfila).ToString().Trim();
                        if (sDesc.Length > 0) oGrid.DataTable.SetValue("Desc_Valida", numfila, sDesc.Remove(sDesc.Length - 1, 1));

                        //Enumera Fila
                        oHeader.SetText(numfila, Convert.ToString(numfila + 1));

                        ((StaticText)oForm.Items.Item("lblTD").Specific).Caption = "Total Documentos : " + TotDocu.ToString();
                        ((StaticText)oForm.Items.Item("lblAC").Specific).Caption = "Total Aceptados : " + TotAcep.ToString();
                        ((StaticText)oForm.Items.Item("lblRC").Specific).Caption = "Total Reclamados : " + TotRecl.ToString();
                        ((StaticText)oForm.Items.Item("lblSL").Specific).Caption = "Total Seleccionados : " + TotSele.ToString();

                    }
                }
                oProgressBar.Value = oProgressBar.Maximum;
            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText("Error ValidarDocumento -> " + e.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("Error ValidarDocumento -> " + e.Message + ", TRACE " + e.StackTrace);
            }
            finally
            {
                oProgressBar.Stop();
                FSBOf._ReleaseCOMObject(oProgressBar);
                FSBOApp.StatusBar.SetText("", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None);
            }
        }

        private string ValidarOC_Xml(String DocEntry, String pCardCode, double pMntTotalXml, DateTime pFchEmisXml, out string ValidaRes)
        {
            string respuesta = "";
            string ValidaResL = "";

            try
            {
                SAPbobsCOM.Recordset orsaux = ((SAPbobsCOM.Recordset)FCmpny.GetBusinessObject(BoObjectTypes.BoRecordset));

                if (GlobalSettings.RunningUnderSQLServer)
                    s = @"SELECT COUNT(*) 'Cant', T2.U_FolioRef, T0.U_FchEmis, T0.U_FchVenc, T0.U_RznSoc, T0.U_MntNeto, T0.U_MntExe, T0.U_MntTotal, T0.U_IVA, T0.U_RUTEmisor, ISNULL(REPLACE(T4.LicTradNum,'.',''),'') as 'RUTOC'
                                                    FROM [@VID_FEXMLCR] T2
                                                    JOIN [@VID_FEXMLC] T0 ON T0.Code = T2.Code
                                                    LEFT JOIN OPOR T3 ON CAST(T3.DocNum as NVARCHAR(20)) = T2.U_FolioRef
                                                    LEFT JOIN OCRD T4 ON T3.CardCode = T4.CardCode
                                                    WHERE T2.Code = '{0}'
                                                    AND T2.U_TpoDocRef = '801'
                                                    GROUP BY T2.U_FolioRef
                                                    , T0.U_FchEmis, T0.U_FchVenc, T0.U_RznSoc, T0.U_MntNeto, T0.U_MntExe, T0.U_MntTotal, T0.U_IVA, T0.U_RUTEmisor,T4.LicTradNum";
                else
                    s = @"SELECT COUNT(*) ""Cant"", T2.""U_FolioRef"", T0.""U_FchEmis"", T0.""U_FchVenc"", T0.""U_RznSoc"", T0.""U_MntNeto"", T0.""U_MntExe"", T0.""U_MntTotal"", T0.""U_IVA"", T0.""U_RUTEmisor"", IFNULL(REPLACE(T4.""LicTradNum"",'.',''),'') as ""RUTOC""
                                                  FROM ""@VID_FEXMLCR"" T2
                                                  JOIN ""@VID_FEXMLC"" T0 ON T0.""Code"" = T2.""Code""
                                                  LEFT JOIN ""OPOR"" T3 ON CAST(T3.""DocNum"" as VARCHAR(20)) = T2.""U_FolioRef""
                                                  LEFT JOIN ""OCRD"" T4 ON T3.""CardCode"" = T4.""CardCode""
                                                 WHERE T2.""Code"" = '{0}'
                                                   AND T2.""U_TpoDocRef"" = '801'
                                                  GROUP BY T2.""U_FolioRef"" 
                                                 , T0.""U_FchEmis"", T0.""U_FchVenc"", T0.""U_RznSoc"", T0.""U_MntNeto"", T0.""U_MntExe"", T0.""U_MntTotal"", T0.""U_IVA"", T0.""U_RUTEmisor"",T4.""LicTradNum"" ";
                s = String.Format(s, DocEntry);
                orsaux.DoQuery(s);
                if (((System.Int32)orsaux.Fields.Item("Cant").Value) == 0)
                {
                    respuesta = "No tiene OC, ";
                    ValidaResL = "NO OC";
                    //respuesta = ValidarOC_Manual(DocEntry, pCardCode, pMntTotalXml, pFchEmisXml, out ValidaResL);
                }
                else if (((System.Int32)orsaux.Fields.Item("Cant").Value) > 1)
                {
                    respuesta = "Tiene mas de una OC, ";
                    ValidaResL = "OC > 1";
                }
                else
                    respuesta = "";

                string RutSN = orsaux.Fields.Item("RUTOC").Value.ToString().Trim();
                var RUTxml = ((System.String)orsaux.Fields.Item("U_RUTEmisor").Value).Trim();
                if (RutSN != RUTxml && RutSN.Length > 0)
                {
                    respuesta = "El RUT del SN no coinciden entre XML y OC SAP , ";
                    ValidaResL = "RUT SN OC";
                }

                if (respuesta == "")//asi filtro que solo sea para el caso que tenga una OC
                {
                    var FchEmisXml = ((System.DateTime)orsaux.Fields.Item("U_FchEmis").Value);
                    var FchvencXml = ((System.DateTime)orsaux.Fields.Item("U_FchVenc").Value);
                    var RznSocXml = ((System.String)orsaux.Fields.Item("U_RznSoc").Value).Trim();
                    var MntNetoXml = ((System.Double)orsaux.Fields.Item("U_MntNeto").Value);
                    var MntExeXml = ((System.Double)orsaux.Fields.Item("U_MntExe").Value);
                    var MntTotalXml = ((System.Double)orsaux.Fields.Item("U_MntTotal").Value);
                    var IVAXml = ((System.Double)orsaux.Fields.Item("U_IVA").Value);
                    var FolioOC = ((System.String)orsaux.Fields.Item("U_FolioRef").Value).Trim();

                    if (GlobalSettings.RunningUnderSQLServer)//Busca CardCode
                        s = @"SELECT CardCode, REPLACE(LicTradNum,'.','') as 'RUT' FROM OCRD WHERE REPLACE(LicTradNum,'.','') = '{0}' AND CardType = 'S' AND frozenFor = 'N'";
                    else
                        s = @"SELECT ""CardCode"", REPLACE(""LicTradNum"",'.','') as ""RUT"" FROM ""OCRD"" WHERE REPLACE(""LicTradNum"",'.','') = '{0}' AND ""CardType"" = 'S' AND ""frozenFor"" = 'N'";
                    s = String.Format(s, RUTxml.Replace(".", ""));
                    orsaux.DoQuery(s);
                    var CardCode = "";
                    if (orsaux.RecordCount == 0)
                    {
                        respuesta = "No se ha encontrado proveedor en el Maestro SN, ";
                        ValidaResL = "NO SN";
                    }
                    else

                        CardCode = ((System.String)orsaux.Fields.Item("CardCode").Value).Trim();


                    if (respuesta == "")//si se encontro el SN
                    {
                        //Busca parametros para validar
                        if (GlobalSettings.RunningUnderSQLServer)
                            s = @"SELECT ISNULL(U_FProv,'Y') 'FProv', ISNULL(U_DiasOC,999) 'DiasOC', ISNULL(U_TipoDif,'M') 'TipoDif', ISNULL(U_DifMon,0) 'DifMon'
                                                            , ISNULL(U_DifPor,0.0) 'DifPor', ISNULL(U_EntMer,'N') 'EntMer' 
                                                        FROM [@VID_FEPARAM] ";
                        else
                            s = @"SELECT IFNULL(""U_FProv"",'Y') ""FProv"", IFNULL(""U_DiasOC"",999) ""DiasOC"", IFNULL(""U_TipoDif"",'M') ""TipoDif"", IFNULL(""U_DifMon"",0) ""DifMon""
                                                            , IFNULL(""U_DifPor"",0.0) ""DifPor"", IFNULL(""U_EntMer"",'N') ""EntMer""
                                                        FROM ""@VID_FEPARAM"" ";
                        orsaux.DoQuery(s);
                        if (orsaux.RecordCount > 0)
                        {
                            var TipoDif = ((System.String)orsaux.Fields.Item("TipoDif").Value).Trim();
                            var DifPor = ((System.Double)orsaux.Fields.Item("DifPor").Value);
                            var DifMon = ((System.Double)orsaux.Fields.Item("DifMon").Value);
                            var DiasOC = ((System.Int32)orsaux.Fields.Item("DiasOC").Value);

                            //Busca datos de la OC
                            if (GlobalSettings.RunningUnderSQLServer)
                                s = @"SELECT T0.DocEntry, T0.DocStatus, T0.CANCELED, T0.Confirmed, T0.DocTotal, T0.VatSum, T0.DocDate, COUNT(*) 'Cant'
                                                        FROM OPOR T0
                                                        JOIN POR1 T1 ON T1.DocEntry = T0.DocEntry
                                                        WHERE CAST(T0.DocNum as NVARCHAR(30)) = '{0}'
                                                        AND T0.CardCode = '{1}'
                                                        GROUP BY T0.DocEntry, T0.DocStatus, T0.CANCELED, T0.Confirmed, T0.DocTotal, T0.VatSum, T0.DocDate";
                            else
                                s = @"SELECT T0.""DocEntry"", T0.""DocStatus"", T0.""CANCELED"", T0.""Confirmed"", T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"", COUNT(*) ""Cant""
                                                        FROM ""OPOR"" T0
                                                        JOIN ""POR1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                                                        WHERE CAST(T0.""DocNum"" as NVARCHAR(30)) = '{0}'
                                                        AND T0.""CardCode"" = '{1}'
                                                        GROUP BY T0.""DocEntry"", T0.""DocStatus"", T0.""CANCELED"", T0.""Confirmed"", T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"" ";
                            s = String.Format(s, FolioOC, CardCode);
                            orsaux.DoQuery(s);

                            if (orsaux.RecordCount == 0)
                            {
                                respuesta = "No se ha encontrado OC en SAP, ";
                                ValidaResL = "NO OC";
                            }
                            else
                            {
                                var OCDocEntry = ((System.Int32)orsaux.Fields.Item("DocEntry").Value);
                                var OCDocStatus = ((System.String)orsaux.Fields.Item("DocStatus").Value).Trim();
                                var OCDocTotal = ((System.Double)orsaux.Fields.Item("DocTotal").Value);
                                var OCVatSum = ((System.Double)orsaux.Fields.Item("VatSum").Value);
                                var OCDocDate = ((System.DateTime)orsaux.Fields.Item("DocDate").Value);
                                var CantLineasOC = ((System.Int32)orsaux.Fields.Item("Cant").Value);
                                var AnuladoOC = ((System.String)orsaux.Fields.Item("CANCELED").Value).Trim();
                                var Autorizada = ((System.String)orsaux.Fields.Item("Confirmed").Value).Trim();

                                if (OCDocStatus != "O")
                                {
                                    respuesta = "La OC en SAP esta Cerrada, ";
                                    ValidaResL = "ESTADO OC";
                                }
                                if (OCDocStatus == "O" && Autorizada == "N")
                                {
                                    respuesta = "La OC en SAP No esta Autorizada, ";
                                    ValidaResL = "ESTADO OC";
                                }
                                if (AnuladoOC == "Y")
                                {
                                    respuesta = "La OC en SAP esta Anulada, ";
                                    ValidaResL = "OC ANULADA";
                                }

                                if (GlobalSettings.RunningUnderSQLServer)
                                    s = @"SELECT COUNT(*) 'Cant'
                                                                    FROM [@VID_FEXMLCD]
                                                                    WHERE Code = '{0}'";
                                else
                                    s = @"SELECT COUNT(*) ""Cant""
                                                                    FROM ""@VID_FEXMLCD""
                                                                    WHERE ""Code"" = '{0}'";
                                s = String.Format(s, DocEntry);
                                orsaux.DoQuery(s);
                                var CantLinFE = ((System.Int32)orsaux.Fields.Item("Cant").Value);


                                //if (CantLineasOC != CantLinFE)
                                //    respuesta = "Cant Lineas no coincide, ";

                                if (OCDocStatus == "O" && AnuladoOC == "N")
                                {
                                    //Valida total OC y total FE
                                    s = ValidarDif(TipoDif, DifPor, DifMon, DiasOC, MntTotalXml, OCDocTotal, DateTime.Now, DateTime.Now, "OC");
                                    if (s != "")
                                    {
                                        respuesta = "Total " + s + ", ";
                                        ValidaResL = TipoDif == "M" ? "$" : "%";
                                    }
                                    else
                                    {
                                        //para validar fecha
                                        s = ValidarDif("D", 0, 0, DiasOC, 0, 0, FchEmisXml, OCDocDate, "OC");
                                        if (s != "")
                                        {
                                            respuesta = "Fecha " + s + ", ";
                                            ValidaResL = "DIAS";
                                        }
                                    }
                                }
                            }
                        }//fin if busca datos de parametros FE
                    }//Fin if respuesta por si encontro CardCode
                }//fin if respuesta por si tiene mas de una OC o no tiene
            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText("Error ValidarOC -> " + e.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("Error ValidarOC -> " + e.Message + ", TRACE " + e.StackTrace);
            }
            ValidaRes = ValidaResL;
            return respuesta;
        }

        private string ValidarEM_Xml(String DocEntry, String pCardCode, double pMntTotalXml, DateTime pFchEmisXml, out string ValidaRes)
        {
            string respuesta = "";
            string ValidaResL = "";
            string sEMCode = oDTParams.GetValue("CodEM", 0).ToString().Trim();

            try
            {
                SAPbobsCOM.Recordset orsaux = ((SAPbobsCOM.Recordset)FCmpny.GetBusinessObject(BoObjectTypes.BoRecordset));
                string sBuscarEMDocNum = oDTParams.GetValue("BuscarEMDocNum", 0).ToString().Trim();
                if (sBuscarEMDocNum == "Y") //buscar EM por DocNum
                {
                    if (GlobalSettings.RunningUnderSQLServer)
                        s = @"SELECT COUNT(*) 'Cant', T2.U_FolioRef, T0.U_FchEmis, T0.U_FchVenc, T0.U_RznSoc, T0.U_MntNeto, T0.U_MntExe, T0.U_MntTotal, T0.U_IVA, T0.U_RUTEmisor, ISNULL(REPLACE(T4.LicTradNum,'.',''),'') as 'RUTOC'
                                                    FROM [@VID_FEXMLCR] T2
                                                    JOIN [@VID_FEXMLC] T0 ON T0.Code = T2.Code
                                                    LEFT JOIN OPDN T3 ON T3.DocNum = T2.U_FolioRef
                                                    LEFT JOIN OCRD T4 ON T3.CardCode = T4.CardCode
                                                    WHERE T2.Code = '{0}'
                                                    AND T2.U_TpoDocRef = '{1}'
                                                    GROUP BY T2.U_FolioRef
                                                    , T0.U_FchEmis, T0.U_FchVenc, T0.U_RznSoc, T0.U_MntNeto, T0.U_MntExe, T0.U_MntTotal, T0.U_IVA, T0.U_RUTEmisor,T4.LicTradNum";
                    else
                        s = @"SELECT COUNT(*) ""Cant"", T2.""U_FolioRef"", T0.""U_FchEmis"", T0.""U_FchVenc"", T0.""U_RznSoc"", T0.""U_MntNeto"", T0.""U_MntExe"", T0.""U_MntTotal"", T0.""U_IVA"", T0.""U_RUTEmisor"", IFNULL(REPLACE(T4.""LicTradNum"",'.',''),'') as ""RUTOC""
                                                  FROM ""@VID_FEXMLCR"" T2
                                                  JOIN ""@VID_FEXMLC"" T0 ON T0.""Code"" = T2.""Code""
                                                  LEFT JOIN ""OPDN"" T3 ON T3.""DocNum"" = T2.""U_FolioRef""
                                                  LEFT JOIN ""OCRD"" T4 ON T3.""CardCode"" = T4.""CardCode""
                                                 WHERE T2.""Code"" = '{0}'
                                                   AND T2.""U_TpoDocRef"" = '{1}'
                                                  GROUP BY T2.""U_FolioRef"" 
                                                 , T0.""U_FchEmis"", T0.""U_FchVenc"", T0.""U_RznSoc"", T0.""U_MntNeto"", T0.""U_MntExe"", T0.""U_MntTotal"", T0.""U_IVA"", T0.""U_RUTEmisor"",T4.""LicTradNum"" ";
                }
                else  //Buscar EM por FolioNum
                {
                    if (GlobalSettings.RunningUnderSQLServer)
                        s = @"SELECT COUNT(*) 'Cant', T2.U_FolioRef, T0.U_FchEmis, T0.U_FchVenc, T0.U_RznSoc, T0.U_MntNeto, T0.U_MntExe, T0.U_MntTotal, T0.U_IVA, T0.U_RUTEmisor, ISNULL(REPLACE(T4.LicTradNum,'.',''),'') as 'RUTOC'
                                                    FROM [@VID_FEXMLCR] T2
                                                    JOIN [@VID_FEXMLC] T0 ON T0.Code = T2.Code
                                                    LEFT JOIN OPDN T3 ON T3.FolioNum = T2.U_FolioRef
                                                    LEFT JOIN OCRD T4 ON T3.CardCode = T4.CardCode
                                                    WHERE T2.Code = '{0}'
                                                    AND T2.U_TpoDocRef = '{1}'
                                                    GROUP BY T2.U_FolioRef
                                                    , T0.U_FchEmis, T0.U_FchVenc, T0.U_RznSoc, T0.U_MntNeto, T0.U_MntExe, T0.U_MntTotal, T0.U_IVA, T0.U_RUTEmisor,T4.LicTradNum";
                    else
                        s = @"SELECT COUNT(*) ""Cant"", T2.""U_FolioRef"", T0.""U_FchEmis"", T0.""U_FchVenc"", T0.""U_RznSoc"", T0.""U_MntNeto"", T0.""U_MntExe"", T0.""U_MntTotal"", T0.""U_IVA"", T0.""U_RUTEmisor"", IFNULL(REPLACE(T4.""LicTradNum"",'.',''),'') as ""RUTOC""
                                                  FROM ""@VID_FEXMLCR"" T2
                                                  JOIN ""@VID_FEXMLC"" T0 ON T0.""Code"" = T2.""Code""
                                                  LEFT JOIN ""OPDN"" T3 ON T3.""FolioNum"" = T2.""U_FolioRef""
                                                  LEFT JOIN ""OCRD"" T4 ON T3.""CardCode"" = T4.""CardCode""
                                                 WHERE T2.""Code"" = '{0}'
                                                   AND T2.""U_TpoDocRef"" = '{1}'
                                                  GROUP BY T2.""U_FolioRef"" 
                                                 , T0.""U_FchEmis"", T0.""U_FchVenc"", T0.""U_RznSoc"", T0.""U_MntNeto"", T0.""U_MntExe"", T0.""U_MntTotal"", T0.""U_IVA"", T0.""U_RUTEmisor"",T4.""LicTradNum"" ";
                }
                    s = String.Format(s, DocEntry, sEMCode);
                orsaux.DoQuery(s);
                if (((System.Int32)orsaux.Fields.Item("Cant").Value) == 0)
                {
                    respuesta = "No tiene EM, ";
                    ValidaResL = "NO EM";
                    //respuesta = ValidarOC_Manual(DocEntry, pCardCode, pMntTotalXml, pFchEmisXml, out ValidaResL);
                }
                else if (((System.Int32)orsaux.Fields.Item("Cant").Value) > 1)
                {
                    respuesta = "Tiene mas de una EM, ";
                    ValidaResL = "EM > 1";
                }
                else
                    respuesta = "";

                string RutSN = orsaux.Fields.Item("RUTOC").Value.ToString().Trim();
                var RUTxml = ((System.String)orsaux.Fields.Item("U_RUTEmisor").Value).Trim();
                if (RutSN != RUTxml && RutSN.Length > 0)
                {
                    respuesta = "El RUT del SN no coinciden entre XML y EM SAP , ";
                    ValidaResL = "RUT SN EM";
                }

                if (respuesta == "")//asi filtro que solo sea para el caso que tenga una OC
                {
                    var FchEmisXml = ((System.DateTime)orsaux.Fields.Item("U_FchEmis").Value);
                    var FchvencXml = ((System.DateTime)orsaux.Fields.Item("U_FchVenc").Value);
                    var RznSocXml = ((System.String)orsaux.Fields.Item("U_RznSoc").Value).Trim();
                    var MntNetoXml = ((System.Double)orsaux.Fields.Item("U_MntNeto").Value);
                    var MntExeXml = ((System.Double)orsaux.Fields.Item("U_MntExe").Value);
                    var MntTotalXml = ((System.Double)orsaux.Fields.Item("U_MntTotal").Value);
                    var IVAXml = ((System.Double)orsaux.Fields.Item("U_IVA").Value);
                    var FolioOC = ((System.String)orsaux.Fields.Item("U_FolioRef").Value).Trim();

                    if (GlobalSettings.RunningUnderSQLServer)//Busca CardCode
                        s = @"SELECT CardCode, REPLACE(LicTradNum,'.','') as 'RUT' FROM OCRD WHERE REPLACE(LicTradNum,'.','') = '{0}' AND CardType = 'S' AND frozenFor = 'N'";
                    else
                        s = @"SELECT ""CardCode"", REPLACE(""LicTradNum"",'.','') as ""RUT"" FROM ""OCRD"" WHERE REPLACE(""LicTradNum"",'.','') = '{0}' AND ""CardType"" = 'S' AND ""frozenFor"" = 'N'";
                    s = String.Format(s, RUTxml.Replace(".", ""));
                    orsaux.DoQuery(s);
                    var CardCode = "";
                    if (orsaux.RecordCount == 0)
                    {
                        respuesta = "No se ha encontrado proveedor en el Maestro SN, ";
                        ValidaResL = "NO SN";
                    }
                    else

                        CardCode = ((System.String)orsaux.Fields.Item("CardCode").Value).Trim();


                    if (respuesta == "")//si se encontro el SN
                    {
                        //Busca parametros para validar
                        if (GlobalSettings.RunningUnderSQLServer)
                            s = @"SELECT ISNULL(U_FProv,'Y') 'FProv', ISNULL(U_DiasOC,999) 'DiasOC', ISNULL(U_TipoDif,'M') 'TipoDif', ISNULL(U_DifMon,0) 'DifMon'
                                                            , ISNULL(U_DifPor,0.0) 'DifPor', ISNULL(U_EntMer,'N') 'EntMer' 
                                                        FROM [@VID_FEPARAM] ";
                        else
                            s = @"SELECT IFNULL(""U_FProv"",'Y') ""FProv"", IFNULL(""U_DiasOC"",999) ""DiasOC"", IFNULL(""U_TipoDif"",'M') ""TipoDif"", IFNULL(""U_DifMon"",0) ""DifMon""
                                                            , IFNULL(""U_DifPor"",0.0) ""DifPor"", IFNULL(""U_EntMer"",'N') ""EntMer""
                                                        FROM ""@VID_FEPARAM"" ";
                        orsaux.DoQuery(s);
                        if (orsaux.RecordCount > 0)
                        {
                            var TipoDif = ((System.String)orsaux.Fields.Item("TipoDif").Value).Trim();
                            var DifPor = ((System.Double)orsaux.Fields.Item("DifPor").Value);
                            var DifMon = ((System.Double)orsaux.Fields.Item("DifMon").Value);
                            var DiasOC = ((System.Int32)orsaux.Fields.Item("DiasOC").Value);
                            //Busca datos de la EM
                            if (sBuscarEMDocNum == "Y") //buscar EM por DocNum
                            {
                                if (GlobalSettings.RunningUnderSQLServer)
                                    s = @"SELECT T0.DocEntry, T0.DocStatus, T0.CANCELED, ,T0.Confirmed, T0.DocTotal, T0.VatSum, T0.DocDate, COUNT(*) 'Cant'
                                                        FROM OPDN T0
                                                        JOIN PDN1 T1 ON T1.DocEntry = T0.DocEntry
                                                        WHERE CAST(T0.DocNum as NVARCHAR(30)) = '{0}'
                                                        AND T0.CardCode = '{1}'
                                                        GROUP BY T0.DocEntry, T0.DocStatus, T0.CANCELED, T0.Confirmed, T0.DocTotal, T0.VatSum, T0.DocDate";
                                else
                                    s = @"SELECT T0.""DocEntry"", T0.""DocStatus"", T0.""CANCELED"",T0.""Confirmed"", T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"", COUNT(*) ""Cant""
                                                        FROM ""OPDN"" T0
                                                        JOIN ""PDN1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                                                        WHERE CAST(T0.""DocNum"" as NVARCHAR(30)) = '{0}'
                                                        AND T0.""CardCode"" = '{1}'
                                                        GROUP BY T0.""DocEntry"", T0.""DocStatus"", T0.""CANCELED"", T0.""Confirmed"", T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"" ";
                            }
                            else
                            {
                                if (GlobalSettings.RunningUnderSQLServer)
                                    s = @"SELECT T0.DocEntry, T0.DocStatus, T0.CANCELED, ,T0.Confirmed, T0.DocTotal, T0.VatSum, T0.DocDate, COUNT(*) 'Cant'
                                                        FROM OPDN T0
                                                        JOIN PDN1 T1 ON T1.DocEntry = T0.DocEntry
                                                        WHERE CAST(T0.FolioNum as NVARCHAR(30)) = '{0}'
                                                        AND T0.CardCode = '{1}'
                                                        GROUP BY T0.DocEntry, T0.DocStatus, T0.CANCELED, T0.Confirmed, T0.DocTotal, T0.VatSum, T0.DocDate";
                                else
                                    s = @"SELECT T0.""DocEntry"", T0.""DocStatus"", T0.""CANCELED"",T0.""Confirmed"", T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"", COUNT(*) ""Cant""
                                                        FROM ""OPDN"" T0
                                                        JOIN ""PDN1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                                                        WHERE CAST(T0.""FolioNum"" as NVARCHAR(30)) = '{0}'
                                                        AND T0.""CardCode"" = '{1}'
                                                        GROUP BY T0.""DocEntry"", T0.""DocStatus"", T0.""CANCELED"", T0.""Confirmed"", T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"" ";

                            }
                            s = String.Format(s, FolioOC, CardCode);
                            orsaux.DoQuery(s);

                            if (orsaux.RecordCount == 0)
                            {
                                respuesta = "No se ha encontrado EM en SAP, ";
                                ValidaResL = "NO EM";
                            }
                            else
                            {
                                var OCDocEntry = ((System.Int32)orsaux.Fields.Item("DocEntry").Value);
                                var OCDocStatus = ((System.String)orsaux.Fields.Item("DocStatus").Value).Trim();
                                var OCDocTotal = ((System.Double)orsaux.Fields.Item("DocTotal").Value);
                                var OCVatSum = ((System.Double)orsaux.Fields.Item("VatSum").Value);
                                var OCDocDate = ((System.DateTime)orsaux.Fields.Item("DocDate").Value);
                                var CantLineasOC = ((System.Int32)orsaux.Fields.Item("Cant").Value);
                                var AnuladoOC = ((System.String)orsaux.Fields.Item("CANCELED").Value).Trim();
                                var Autorizada = ((System.String)orsaux.Fields.Item("Confirmed").Value).Trim();


                                if (OCDocStatus != "O")
                                {
                                    respuesta = "La EM en SAP esta Cerrada, ";
                                    ValidaResL = "ESTADO EM";
                                }
                                if (OCDocStatus == "O" && Autorizada == "N")
                                {
                                    respuesta = "La EM en SAP No esta Autorizada, ";
                                    ValidaResL = "ESTADO EM";
                                }
                                if (AnuladoOC == "Y")
                                {
                                    respuesta = "La EM en SAP esta Anulada, ";
                                    ValidaResL = "EM ANULADA";
                                }

                                if (GlobalSettings.RunningUnderSQLServer)
                                    s = @"SELECT COUNT(*) 'Cant'
                                                                    FROM [@VID_FEXMLCD]
                                                                    WHERE Code = '{0}'";
                                else
                                    s = @"SELECT COUNT(*) ""Cant""
                                                                    FROM ""@VID_FEXMLCD""
                                                                    WHERE ""Code"" = '{0}'";
                                s = String.Format(s, DocEntry);
                                orsaux.DoQuery(s);
                                var CantLinFE = ((System.Int32)orsaux.Fields.Item("Cant").Value);


                                //if (CantLineasOC != CantLinFE)
                                //    respuesta = "Cant Lineas no coincide, ";

                                if (OCDocStatus == "O" && AnuladoOC == "N")
                                {
                                    //Valida total OC y total FE
                                    s = ValidarDif(TipoDif, DifPor, DifMon, DiasOC, MntTotalXml, OCDocTotal, DateTime.Now, DateTime.Now, "EM");
                                    if (s != "")
                                    {
                                        respuesta = "Total " + s + ", ";
                                        ValidaResL = TipoDif == "M" ? "$" : "%";
                                    }
                                    else
                                    {
                                        //para validar fecha
                                        s = ValidarDif("D", 0, 0, DiasOC, 0, 0, FchEmisXml, OCDocDate, "EM");
                                        if (s != "")
                                        {
                                            respuesta = "Fecha " + s + ", ";
                                            ValidaResL = "DIAS";
                                        }
                                    }
                                }
                            }
                        }//fin if busca datos de parametros FE
                    }//Fin if respuesta por si encontro CardCode
                }//fin if respuesta por si tiene mas de una OC o no tiene
            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText("Error ValidarEM -> " + e.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("Error ValidarEM -> " + e.Message + ", TRACE " + e.StackTrace);
            }
            ValidaRes = ValidaResL;
            return respuesta;
        }

        private string ValidarOC_Manual(String DocNum, String CardCode, double MntTotalXml, DateTime FchEmisXml, out string ValidaRes)
        {
            string respuesta = "";
            string ValidaResL = "";

            try
            {
                SAPbobsCOM.Recordset orsaux = ((SAPbobsCOM.Recordset)FCmpny.GetBusinessObject(BoObjectTypes.BoRecordset));
                if (respuesta == "")
                {
                    //Busca parametros para validar
                    if (GlobalSettings.RunningUnderSQLServer)
                        s = @"SELECT ISNULL(U_FProv,'Y') 'FProv', ISNULL(U_DiasOC,999) 'DiasOC', ISNULL(U_TipoDif,'M') 'TipoDif', ISNULL(U_DifMon,0) 'DifMon'
                                                            , ISNULL(U_DifPor,0.0) 'DifPor', ISNULL(U_EntMer,'N') 'EntMer' 
                                                        FROM [@VID_FEPARAM] ";
                    else
                        s = @"SELECT IFNULL(""U_FProv"",'Y') ""FProv"", IFNULL(""U_DiasOC"",999) ""DiasOC"", IFNULL(""U_TipoDif"",'M') ""TipoDif"", IFNULL(""U_DifMon"",0) ""DifMon""
                                                            , IFNULL(""U_DifPor"",0.0) ""DifPor"", IFNULL(""U_EntMer"",'N') ""EntMer""
                                                        FROM ""@VID_FEPARAM"" ";
                    orsaux.DoQuery(s);
                    if (orsaux.RecordCount > 0)
                    {
                        var TipoDif = ((System.String)orsaux.Fields.Item("TipoDif").Value).Trim();
                        var DifPor = ((System.Double)orsaux.Fields.Item("DifPor").Value);
                        var DifMon = ((System.Double)orsaux.Fields.Item("DifMon").Value);
                        var DiasOC = ((System.Int32)orsaux.Fields.Item("DiasOC").Value);

                        //Busca datos de la OC
                        if (GlobalSettings.RunningUnderSQLServer)
                            s = @"SELECT T0.DocEntry, T0.CardCode, T0.DocStatus, T0.Confirmed, T0.CANCELED, T0.DocTotal, T0.VatSum, T0.DocDate, COUNT(*) 'Cant'
                                                        FROM OPOR T0
                                                        JOIN POR1 T1 ON T1.DocEntry = T0.DocEntry
                                                        WHERE CAST(T0.DocNum as NVARCHAR(30)) = '{0}'
                                                        GROUP BY T0.DocEntry, T0.CardCode, T0.DocStatus,  T0.CANCELED, T0.Confirmed, T0.DocTotal, T0.VatSum, T0.DocDate";
                        else
                            s = @"SELECT T0.""DocEntry"", T0.""CardCode"", T0.""DocStatus"", T0.""Confirmed"", T0.""CANCELED"",  T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"", COUNT(*) ""Cant""
                                                        FROM ""OPOR"" T0
                                                        JOIN ""POR1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                                                        WHERE CAST(T0.""DocNum"" as NVARCHAR(30)) = '{0}'
                                                        GROUP BY T0.""DocEntry"", T0.""CardCode"", T0.""DocStatus"",  T0.""CANCELED"",  T0.""Confirmed"",T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"" ";
                        s = String.Format(s, DocNum);
                        orsaux.DoQuery(s);

                        if (orsaux.RecordCount == 0)
                        {
                            respuesta = "No se ha encontrado Orden de Compra en SAP, ";
                            ValidaResL = "NO OC";
                        }
                        else
                        {
                            if (CardCode != orsaux.Fields.Item("CardCode").Value.ToString().Trim())
                            {
                                respuesta = "El RUT del SN no coinciden entre XML y OC SAP , ";
                                ValidaResL = "RUT SN OC";
                            }
                            else
                            {
                                var OCDocEntry = ((System.Int32)orsaux.Fields.Item("DocEntry").Value);
                                var OCDocStatus = ((System.String)orsaux.Fields.Item("DocStatus").Value).Trim();
                                var OCDocTotal = ((System.Double)orsaux.Fields.Item("DocTotal").Value);
                                var OCVatSum = ((System.Double)orsaux.Fields.Item("VatSum").Value);
                                var OCDocDate = ((System.DateTime)orsaux.Fields.Item("DocDate").Value);
                                var CantLineasOC = ((System.Int32)orsaux.Fields.Item("Cant").Value);
                                var AnuladoOC = ((System.String)orsaux.Fields.Item("CANCELED").Value).Trim();
                                var Autorizada = ((System.String)orsaux.Fields.Item("Confirmed").Value).Trim();

                                if (OCDocStatus != "O")
                                {
                                    respuesta = "La OC en SAP esta Cerrada, ";
                                    ValidaResL = "ESTADO OC";
                                }
                                if (OCDocStatus == "O" && Autorizada == "N")
                                {
                                    respuesta = "La OC en SAP No esta Autorizada, ";
                                    ValidaResL = "ESTADO OC";
                                }
                                if (AnuladoOC == "Y")
                                {
                                    respuesta = "La OC en SAP esta Anulada, ";
                                    ValidaResL = "OC ANULADA";
                                }

                                if (OCDocStatus == "O" && AnuladoOC == "N")
                                {
                                    //Valida total OC y total FE
                                    s = ValidarDif(TipoDif, DifPor, DifMon, DiasOC, MntTotalXml, OCDocTotal, DateTime.Now, DateTime.Now, "OC");
                                    if (s != "")
                                    {
                                        respuesta = "Total " + s + ", ";
                                        ValidaResL = TipoDif == "M" ? "$" : "%";
                                    }
                                    else
                                    {
                                        //para validar fecha
                                        s = ValidarDif("D", 0, 0, DiasOC, 0, 0, FchEmisXml, OCDocDate, "OC");
                                        if (s != "")
                                        {
                                            respuesta = "Fecha " + s + ", ";
                                            ValidaResL = "DIAS";
                                        }
                                    }
                                }
                            }
                        }
                    }//fin if busca datos de parametros FE
                }//Fin if respuesta por si encontro CardCode
            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText("Error ValidarOC -> " + e.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("Error ValidarOC -> " + e.Message + ", TRACE " + e.StackTrace);
            }
            ValidaRes = ValidaResL;
            return respuesta;
        }

        private string ValidarEM_Manual(String DocNum, String CardCode, double MntTotalXml, DateTime FchEmisXml, out string ValidaRes)
        {
            string respuesta = "";
            string ValidaResL = "";

            try
            {
                SAPbobsCOM.Recordset orsaux = ((SAPbobsCOM.Recordset)FCmpny.GetBusinessObject(BoObjectTypes.BoRecordset));
                if (respuesta == "")
                {
                    //Busca parametros para validar
                    if (GlobalSettings.RunningUnderSQLServer)
                        s = @"SELECT ISNULL(U_FProv,'Y') 'FProv', ISNULL(U_DiasOC,999) 'DiasOC', ISNULL(U_TipoDif,'M') 'TipoDif', ISNULL(U_DifMon,0) 'DifMon'
                                                            , ISNULL(U_DifPor,0.0) 'DifPor', ISNULL(U_EntMer,'N') 'EntMer' 
                                                        FROM [@VID_FEPARAM] ";
                    else
                        s = @"SELECT IFNULL(""U_FProv"",'Y') ""FProv"", IFNULL(""U_DiasOC"",999) ""DiasOC"", IFNULL(""U_TipoDif"",'M') ""TipoDif"", IFNULL(""U_DifMon"",0) ""DifMon""
                                                            , IFNULL(""U_DifPor"",0.0) ""DifPor"", IFNULL(""U_EntMer"",'N') ""EntMer""
                                                        FROM ""@VID_FEPARAM"" ";
                    orsaux.DoQuery(s);
                    if (orsaux.RecordCount > 0)
                    {
                        var TipoDif = ((System.String)orsaux.Fields.Item("TipoDif").Value).Trim();
                        var DifPor = ((System.Double)orsaux.Fields.Item("DifPor").Value);
                        var DifMon = ((System.Double)orsaux.Fields.Item("DifMon").Value);
                        var DiasOC = ((System.Int32)orsaux.Fields.Item("DiasOC").Value);

                        //Busca datos de la EM
                        string sBuscarEMDocNum = oDTParams.GetValue("BuscarEMDocNum", 0).ToString().Trim();
                        if (sBuscarEMDocNum == "Y") //buscar EM por DocNum
                        {
                            if (GlobalSettings.RunningUnderSQLServer)
                                s = @"SELECT T0.DocEntry, T0.CardCode, T0.DocStatus, T0.Confirmed, T0.CANCELED, T0.DocTotal, T0.VatSum, T0.DocDate, COUNT(*) 'Cant'
                                                        FROM OPDN T0
                                                        JOIN PDN1 T1 ON T1.DocEntry = T0.DocEntry
                                                        WHERE CAST(T0.DocNum as NVARCHAR(30)) = '{0}'
                                                        GROUP BY T0.DocEntry, T0.CardCode, T0.DocStatus,  T0.CANCELED , T0.Confirmed, T0.DocTotal, T0.VatSum, T0.DocDate";
                            else
                                s = @"SELECT T0.""DocEntry"", T0.""CardCode"", T0.""DocStatus"", T0.""Confirmed"", T0.""CANCELED"", T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"", COUNT(*) ""Cant""
                                                        FROM ""OPDN"" T0
                                                        JOIN ""PDN1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                                                        WHERE CAST(T0.""DocNum"" as NVARCHAR(30)) = '{0}'
                                                        GROUP BY T0.""DocEntry"", T0.""CardCode"", T0.""DocStatus"",  T0.""CANCELED"", T0.""Confirmed"", T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"" ";
                        }
                        else //Buscar EM por FolioNum
                        {
                            if (GlobalSettings.RunningUnderSQLServer)
                                s = @"SELECT T0.DocEntry, T0.CardCode, T0.DocStatus, T0.Confirmed, T0.CANCELED, T0.DocTotal, T0.VatSum, T0.DocDate, COUNT(*) 'Cant'
                                                        FROM OPDN T0
                                                        JOIN PDN1 T1 ON T1.DocEntry = T0.DocEntry
                                                        WHERE CAST(T0.FolioNum as NVARCHAR(30)) = '{0}'
                                                        GROUP BY T0.DocEntry, T0.CardCode, T0.DocStatus,  T0.CANCELED , T0.Confirmed, T0.DocTotal, T0.VatSum, T0.DocDate";
                            else
                                s = @"SELECT T0.""DocEntry"", T0.""CardCode"", T0.""DocStatus"", T0.""Confirmed"", T0.""CANCELED"", T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"", COUNT(*) ""Cant""
                                                        FROM ""OPDN"" T0
                                                        JOIN ""PDN1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                                                        WHERE CAST(T0.""FolioNum"" as NVARCHAR(30)) = '{0}'
                                                        GROUP BY T0.""DocEntry"", T0.""CardCode"", T0.""DocStatus"",  T0.""CANCELED"", T0.""Confirmed"", T0.""DocTotal"", T0.""VatSum"", T0.""DocDate"" ";

                        }
                        s = String.Format(s, DocNum);
                        orsaux.DoQuery(s);

                        if (orsaux.RecordCount == 0)
                        {
                            respuesta = "No se ha encontrado la Entrada de Mercancia en SAP, ";
                            ValidaResL = "NO EM";
                        }
                        else
                        {
                            if (CardCode != orsaux.Fields.Item("CardCode").Value.ToString().Trim())
                            {
                                respuesta = "El RUT del SN no coinciden entre XML y EM SAP , ";
                                ValidaResL = "RUT SN EM";
                            }
                            else
                            {
                                var OCDocEntry = ((System.Int32)orsaux.Fields.Item("DocEntry").Value);
                                var OCDocStatus = ((System.String)orsaux.Fields.Item("DocStatus").Value).Trim();
                                var OCDocTotal = ((System.Double)orsaux.Fields.Item("DocTotal").Value);
                                var OCVatSum = ((System.Double)orsaux.Fields.Item("VatSum").Value);
                                var OCDocDate = ((System.DateTime)orsaux.Fields.Item("DocDate").Value);
                                var CantLineasOC = ((System.Int32)orsaux.Fields.Item("Cant").Value);
                                var AnuladoOC = ((System.String)orsaux.Fields.Item("CANCELED").Value).Trim();
                                var Autorizada = ((System.String)orsaux.Fields.Item("Confirmed").Value).Trim();

                                if (OCDocStatus != "O")
                                {
                                    respuesta = "La EM en SAP esta Cerrada, ";
                                    ValidaResL = "ESTADO EM";
                                }
                                if (OCDocStatus == "O" && Autorizada == "N")
                                {
                                    respuesta = "La EM en SAP No esta Autorizada, ";
                                    ValidaResL = "ESTADO EM";
                                }
                                if (AnuladoOC == "Y")
                                {
                                    respuesta = "La EM en SAP esta Anulada, ";
                                    ValidaResL = "EM ANULADA";
                                }

                                if (OCDocStatus == "O" && AnuladoOC == "N")
                                {
                                    //Valida total OC y total FE
                                    s = ValidarDif(TipoDif, DifPor, DifMon, DiasOC, MntTotalXml, OCDocTotal, DateTime.Now, DateTime.Now, "EM");
                                    if (s != "")
                                    {
                                        respuesta = "Total " + s + ", ";
                                        ValidaResL = TipoDif == "M" ? "$" : "%";
                                    }
                                    else
                                    {
                                        //para validar fecha
                                        s = ValidarDif("D", 0, 0, DiasOC, 0, 0, FchEmisXml, OCDocDate, "EM");
                                        if (s != "")
                                        {
                                            respuesta = "Fecha " + s + ", ";
                                            ValidaResL = "DIAS";
                                        }
                                    }
                                }
                            }
                        }
                    }//fin if busca datos de parametros FE
                }//Fin if respuesta por si encontro CardCode
            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText("Error ValidarEM -> " + e.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("Error ValidarEM -> " + e.Message + ", TRACE " + e.StackTrace);
            }
            ValidaRes = ValidaResL;
            return respuesta;
        }

        private String ValidarDif(String TipoDif, Double DifPor, Double DifMon, Int32 Dias, Double Monto, Double MontoOC, DateTime Fecha, DateTime FechaOC, string TipoDoc)
        {
            Double DifCal;
            try
            {
                if (TipoDif == "M")  //revisar pra que sea + o - 
                {
                    DifCal = MontoOC - Monto;
                    if (DifCal < 0)
                        DifCal = DifCal * -1;
                    if (DifCal > DifMon)
                        return "Documento presenta diferencia en valor " + TipoDoc;
                }
                else if (TipoDif == "P")   //revisar pra que sea + o -
                {
                    DifCal = MontoOC - Monto;
                    if (DifCal < 0)
                        DifCal = DifCal * -1;

                    var Valor = (DifPor * MontoOC) / 100;

                    if (DifCal > Valor)
                        return "Documento presenta diferencia en valor " + TipoDoc;
                }

                //Para las diferencias de dias con la OC
                if (TipoDif == "D")
                {
                    var diferenciaDias = Fecha - FechaOC;
                    var difReal = diferenciaDias.Days;
                    if (difReal < 0)
                        difReal = difReal * -1;
                    if (difReal > Dias)
                        return "Se supera los dias maximo entre " + TipoDoc + " y FE";
                }
                return "";
            }
            catch (Exception e)
            {
                return "Error Validar";
            }
        }

    }//fin class
}
