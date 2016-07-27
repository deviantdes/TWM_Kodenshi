using System;
using System.Collections.Generic;
using System.Text;


namespace TWM_KDS_AddOn
{
    [Form("twmKDS_EXPXML", true, "Export to XML", SBOAddon.gcAddOnName, 3)]
    //[Authorization("twmOPS_DBTST", "DB Trans Setup", SBOAddon.gcAddOnName, SAPbobsCOM.BoUPTOptions.bou_FullReadNone)]
    public class twmKDS_EXPXML
    {
        SAPbouiCOM.Form _oForm = null;
        SAPbouiCOM.CheckBox _cbARINV = null;
        SAPbouiCOM.CheckBox _cbARCN = null;
        SAPbouiCOM.CheckBox _cbAPINV = null;
        SAPbouiCOM.CheckBox _cbAPCN = null;
        SAPbouiCOM.EditText _txtDTFRM = null;
        SAPbouiCOM.EditText _txtDTTO = null;
        SAPbouiCOM.Button _btnView = null;
        SAPbouiCOM.Button _btnExport = null;
        SAPbouiCOM.Button _btnCP = null;
        SAPbouiCOM.EditText _txtPath = null;

        SAPbouiCOM.Grid _grid_Trans = null;
        SAPbouiCOM.DataTable _dt_grid_Trans = null;
        private bool toggleGridCheckBox = false;
        Dictionary<String, String> branchCheckDict = new Dictionary<String, String>();

        public twmKDS_EXPXML()
        {
            try
            {
                //Draw the form
                String sFileName = string.Format("{0}.xml", this.GetType().Name);
                bool bSuccess = true;
                if (System.IO.File.Exists(Environment.CurrentDirectory + "\\" + sFileName))
                {
                    System.Xml.XmlDocument oXml = new System.Xml.XmlDocument();
                    oXml.Load(Environment.CurrentDirectory + "\\" + sFileName);
                    String sXml = eCommon.ModifySize(oXml.InnerXml);
                    try
                    {
                        eCommon.SBO_Application.LoadBatchActions(ref sXml);
                    }
                    catch
                    {
                        bSuccess = false;
                    }
                }
                else
                {
                    String ResourceName = string.Format("{0}.Src.Resource.{1}", System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, sFileName);

                    String sXml = eCommon.ModifySize(eCommon.GetXMLResource(ResourceName));
                    try
                    {
                        eCommon.SBO_Application.LoadBatchActions(ref sXml);
                    }
                    catch {
                        bSuccess = false;
                    }
                }

                if (bSuccess)
                {
                    _oForm = eCommon.SBO_Application.Forms.ActiveForm;
                    if (_oForm.TypeEx.StartsWith("-"))
                    {
                        //a UDF form is opened. Close it.
                        String UDFFormUID = _oForm.UniqueID;
                        String ParentFormUID = eCommon.GetParentFormUID(_oForm);
                        _oForm.Close();

                        _oForm = eCommon.SBO_Application.Forms.ActiveForm;
                        if (_oForm.UniqueID != ParentFormUID)
                            _oForm = eCommon.SBO_Application.Forms.Item(ParentFormUID);
                    }
                    _oForm.EnableMenu("6913", false);

                    GetItemReferences();
                    // Initform can be use if you want to fill anything or color anything
                    InitForm();
                    if (!SBOAddon.oOpenForms.Contains(_oForm.UniqueID))
                        SBOAddon.oOpenForms.Add(_oForm.UniqueID, this);

                    _oForm.Visible = true;
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        public twmKDS_EXPXML(SAPbouiCOM.Form oForm)
        {
            _oForm = oForm;
            GetItemReferences();
            if (!SBOAddon.oOpenForms.Contains(_oForm.UniqueID))
                SBOAddon.oOpenForms.Add(_oForm.UniqueID, this);
        }

        public twmKDS_EXPXML(String FormUID)
        {
            _oForm = eCommon.SBO_Application.Forms.Item(FormUID);
            GetItemReferences();
            if (!SBOAddon.oOpenForms.Contains(_oForm.UniqueID))
                SBOAddon.oOpenForms.Add(_oForm.UniqueID, this);

        }

        private void InitForm()
        {
            try
            {
                _oForm.Freeze(true);
                _grid_Trans.DataTable = _dt_grid_Trans;
                ((SAPbouiCOM.EditText)_oForm.Items.Item("txtDTFRM").Specific).Active = true;  

                // Get path from DB settings
                _txtPath.Value = SBOAddon_DB.Settings_xml_Path;

            }
            catch (Exception Ex)
            {
                eCommon.SBO_Application.MessageBox(Ex.Message);
            }
            finally
            {
                _oForm.Freeze(false);
            }
        }

        public void GetItemReferences()
        {
            try
            {
                _oForm.Freeze(false);
                _cbARINV = _oForm.Items.Item("cbARINV").Specific as SAPbouiCOM.CheckBox;
                _cbARCN = _oForm.Items.Item("cbARCN").Specific as SAPbouiCOM.CheckBox;
                _cbAPINV = _oForm.Items.Item("cbAPINV").Specific as SAPbouiCOM.CheckBox;
                _cbAPCN = _oForm.Items.Item("cbAPCN").Specific as SAPbouiCOM.CheckBox;
                _txtDTFRM = _oForm.Items.Item("txtDTFRM").Specific as SAPbouiCOM.EditText;
                _txtDTTO = _oForm.Items.Item("txtDTTO").Specific as SAPbouiCOM.EditText;
                _btnView = _oForm.Items.Item("btnView").Specific as SAPbouiCOM.Button;
                _btnCP = _oForm.Items.Item("btnCP").Specific as SAPbouiCOM.Button;
                _txtPath = _oForm.Items.Item("txtPath").Specific as SAPbouiCOM.EditText;
                _btnExport = _oForm.Items.Item("btnExport").Specific as SAPbouiCOM.Button;
                _grid_Trans = _oForm.Items.Item("grid_Trans").Specific as SAPbouiCOM.Grid;
                _dt_grid_Trans = _oForm.DataSources.DataTables.Item("grid_Trans");

                _btnView.PressedAfter += _btnView_PressedAfter;
                _btnExport.PressedAfter += _btnExport_PressedAfter;
                // Double click events
                _grid_Trans.DoubleClickAfter += _grid_Trans_DoubleClickAfter;
                _btnCP.PressedAfter += _btnCP_PressedAfter;
      
            }
            catch (Exception Ex){eCommon.SBO_Application.MessageBox(Ex.Message);}
        }

        void _btnCP_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {

            SBOCustom.MyFileDialog fileDiag = new SBOCustom.MyFileDialog(eCommon.SBO_Application);
            _txtPath.Value = fileDiag.OpenFolderDialog() + "\\";
        }

        void _grid_Trans_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //SAPbouiCOM.ProgressBar oProgressBar = null;
            try
            {
                if (pVal.Row == -1 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_CTRL)
                    return;

                if (pVal.ColUID == "Checked" && pVal.Row == -1)
                    checkAllGrid(_grid_Trans);

                else if (pVal.ColUID == "Checked" && pVal.Row > -1)
                {
                    // if it is collaspeable
                    if (!_grid_Trans.Rows.IsLeaf(pVal.Row))
                    {
                        //in case there are more branch and not leafs, loop to the last branch
                        int _Leaf = 0;
                        for (int i = 1; ; i++)
                            if (_grid_Trans.Rows.IsLeaf(pVal.Row + i))
                            {
                                _Leaf = i;
                                break;
                            }

                        //use GetDataTableRowIndex when grid is collapse to get true row
                        String objType = _grid_Trans.DataTable.GetValue("ObjType", _grid_Trans.GetDataTableRowIndex(pVal.Row + _Leaf)).ToString();
                        String check = _grid_Trans.DataTable.GetValue("Checked", _grid_Trans.GetDataTableRowIndex(pVal.Row + _Leaf)).ToString();

                        if (check == "Y")
                            check = "N";
                        else
                            check = "Y";

                        _oForm.Freeze(true);
                       
                        /*
                        try
                        {
                            if (oProgressBar==null)
                            oProgressBar = eCommon.SBO_Application.StatusBar.CreateProgressBar("Export in progress", 100, false);
                        }
                        catch { }
                        */

                        // get all the rows with this objType
                        int[] Rows = eCommon.DataTableIndexOf(_dt_grid_Trans, "ObjType", objType);
                        /* int _progress = 0;

                        if (oProgressBar != null)
                        {
                            oProgressBar.Text = "Export in progress...";
                            oProgressBar.Maximum = Rows.Length;
                        }
                        */
                        foreach (int Row in Rows)
                        {
                            _grid_Trans.DataTable.SetValue("Checked", Row, check);
                            /*
                            _progress += 1;
                            
                            if (oProgressBar != null)
                                oProgressBar.Value = _progress;
                             */
                        }

                        // Set this objType to checked/uncheck and call populate grid again
                        //branchCheckDict[objType] = check;
                        //populateGrid();
                        _oForm.Freeze(false);
                    }
                }
            }
            catch (Exception ex) { eCommon.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning); }
            finally
            {
                /*
                if (oProgressBar != null)
                {
                    oProgressBar.Stop();
                    oProgressBar = null;
                }
                 */
            }
        }

        // toggle all checkbox in grid
        private void checkAllGrid(SAPbouiCOM.Grid theGrid)
        {
            _oForm.Freeze(true);

            if (!toggleGridCheckBox)
            {
                populateGrid(check: "Y");
                toggleGridCheckBox = true;
            }
            else
            {
                populateGrid(check: "N");
                toggleGridCheckBox = false;
            }

            _oForm.Freeze(false);
        }

        // Export button pressed
        void _btnExport_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (_txtPath.Value.Length < 1)
                {
                    eCommon.SBO_Application.StatusBar.SetText("Invalid Path to save.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return;
                }

                int[] SelectedRows = eCommon.DataTableIndexOf(_dt_grid_Trans, "Checked", "Y");
                if (SelectedRows == null)
                    eCommon.SBO_Application.StatusBar.SetText("No Transactions selected.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                else
                {

                    String str_directory = Environment.CurrentDirectory.ToString();
                    String parent = System.IO.Directory.GetParent(System.IO.Directory.GetParent(System.IO.Directory.GetParent(str_directory).FullName).FullName).FullName;
                    exportXML(SelectedRows);
                }
            }
            catch (Exception ex)
            {
                eCommon.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        // Create XML
        public void exportXML(int[] SelectedRows)
        {
            SAPbobsCOM.Documents oDoc = null;
            SAPbouiCOM.ProgressBar oProgressBar = null;
            try
            {
                _btnExport.Item.Enabled = false;
                String tempName = "";
                eCommon.oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode;
                
                int _progress = 0;
                oProgressBar = eCommon.SBO_Application.StatusBar.CreateProgressBar("Export in progress", SelectedRows.Length, false);
                oProgressBar.Text = "Export in progress...";

                foreach (int rows in SelectedRows)
                {
                    //Sales Invoice
                    if (_grid_Trans.DataTable.GetValue("ObjType", rows).ToString() == "13")
                    {
                        oDoc = eCommon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices) as SAPbobsCOM.Documents;
                        tempName = "ARINVOICE";
                    }
                    //Sales Credit Note
                    else if (_grid_Trans.DataTable.GetValue("ObjType", rows).ToString() == "14")
                    {
                        oDoc = eCommon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes) as SAPbobsCOM.Documents;
                        tempName = "ARCREDITNOTE";
                    }
                    //Purchase Invoice
                    else if (_grid_Trans.DataTable.GetValue("ObjType", rows).ToString() == "18")
                    {
                        oDoc = eCommon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices) as SAPbobsCOM.Documents;
                        tempName = "APINVOICE";
                    }
                    //Purchase Credit Note
                    else if (_grid_Trans.DataTable.GetValue("ObjType", rows).ToString() == "19")
                    {
                        oDoc = eCommon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes) as SAPbobsCOM.Documents;
                        tempName = "APCREDITNOTE";
                    }

                    String temp = _grid_Trans.DataTable.GetValue("DocEntry", rows).ToString();
                    oDoc.GetByKey(int.Parse(temp));

                    String filepath = String.Format(_txtPath.Value + tempName + "_{0}.xml", oDoc.DocNum);

                    oDoc.SaveXML(filepath);

                    _progress += 1;
                    oProgressBar.Value = _progress;
                }


                populateGrid(check: "N");
                _btnExport.Item.Enabled = true;
                eCommon.SBO_Application.StatusBar.SetText("Export Completed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex) { eCommon.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning); }
            finally
            {
                if (oProgressBar != null)
                {
                    oProgressBar.Stop();
                    oProgressBar = null;
                }
                if (oDoc != null)
                    eCommon.ReleaseComObject(oDoc);
            }
        }

        void _btnView_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (!_cbARINV.Checked && !_cbARCN.Checked && !_cbAPINV.Checked && !_cbAPCN.Checked)
                    eCommon.SBO_Application.StatusBar.SetText("No Transaction type selected.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                else if (_txtDTFRM.Value.Length < 1 || _txtDTTO.Value.Length < 1)
                    eCommon.SBO_Application.StatusBar.SetText("Please specify date range.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                else
                {
                    populateGrid(check:"N");
                }
            }
            catch(Exception ex){eCommon.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);}   
        }

        private void populateGrid(String check = null)
        {
            StringBuilder query = new StringBuilder();

            if (check != null)
            {
                // Set the default check settings
                branchCheckDict = new Dictionary<String, String>();

                if (_cbARINV.Checked)
                {
                    query.Append(String.Format(" UNION SELECT ObjType, '{2}' Checked, DocEntry, DocNum, DocDate, CardCode, CardName, NumAtCard, DocCur, CASE DocTotalFC when 0 then DocTotal else DocTotalFC END as DocTotal FROM OINV WHERE CANCELED ='N' AND  DocDate between '{0}' AND '{1}'", _txtDTFRM.Value, _txtDTTO.Value, check));
                    branchCheckDict.Add("13", check);
                }
                if (_cbARCN.Checked)
                {
                    query.Append(String.Format(" UNION SELECT ObjType, '{2}' Checked, DocEntry, DocNum, DocDate, CardCode, CardName, NumAtCard, DocCur, CASE DocTotalFC when 0 then DocTotal else DocTotalFC END as DocTotal FROM ORIN WHERE CANCELED ='N' AND  DocDate between '{0}' AND '{1}'", _txtDTFRM.Value, _txtDTTO.Value, check));
                    branchCheckDict.Add("14", check);
                }
                if (_cbAPINV.Checked)
                {
                    query.Append(String.Format(" UNION SELECT ObjType, '{2}' Checked, DocEntry, DocNum, DocDate, CardCode, CardName, NumAtCard, DocCur, CASE DocTotalFC when 0 then DocTotal else DocTotalFC END as DocTotal FROM OPCH WHERE CANCELED ='N' AND  DocDate between '{0}' AND '{1}'", _txtDTFRM.Value, _txtDTTO.Value, check));
                    branchCheckDict.Add("18", check);
                }
                if (_cbAPCN.Checked)
                {
                    query.Append(String.Format(" UNION SELECT ObjType, '{2}' Checked, DocEntry, DocNum, DocDate, CardCode, CardName, NumAtCard, DocCur, CASE DocTotalFC when 0 then DocTotal else DocTotalFC END as DocTotal FROM ORPC WHERE CANCELED ='N' AND  DocDate between '{0}' AND '{1}'", _txtDTFRM.Value, _txtDTTO.Value, check));
                    branchCheckDict.Add("19", check);
                }
            }
            // currently this condition is not being used. line 211-212
            else
            {
                if (branchCheckDict.ContainsKey("13"))
                    query.Append(String.Format(" UNION SELECT ObjType, '{2}' Checked, DocEntry, DocNum, DocDate, CardCode, CardName, NumAtCard, DocCur, CASE DocTotalFC when 0 then DocTotal else DocTotalFC END as DocTotal FROM OINV WHERE CANCELED ='N' AND  DocDate between '{0}' AND '{1}'", _txtDTFRM.Value, _txtDTTO.Value, branchCheckDict["13"]));
                if (branchCheckDict.ContainsKey("14"))
                    query.Append(String.Format(" UNION SELECT ObjType, '{2}' Checked, DocEntry, DocNum, DocDate, CardCode, CardName, NumAtCard, DocCur, CASE DocTotalFC when 0 then DocTotal else DocTotalFC END as DocTotal FROM ORIN WHERE CANCELED ='N' AND  DocDate between '{0}' AND '{1}'", _txtDTFRM.Value, _txtDTTO.Value, branchCheckDict["14"]));
                if (branchCheckDict.ContainsKey("18"))
                    query.Append(String.Format(" UNION SELECT ObjType, '{2}' Checked, DocEntry, DocNum, DocDate, CardCode, CardName, NumAtCard, DocCur, CASE DocTotalFC when 0 then DocTotal else DocTotalFC END as DocTotal FROM OPCH WHERE CANCELED ='N' AND  DocDate between '{0}' AND '{1}'", _txtDTFRM.Value, _txtDTTO.Value, branchCheckDict["18"]));
                if (branchCheckDict.ContainsKey("19"))
                    query.Append(String.Format(" UNION SELECT ObjType, '{2}' Checked, DocEntry, DocNum, DocDate, CardCode, CardName, NumAtCard, DocCur, CASE DocTotalFC when 0 then DocTotal else DocTotalFC END as DocTotal FROM ORPC WHERE CANCELED ='N' AND  DocDate between '{0}' AND '{1}'", _txtDTFRM.Value, _txtDTTO.Value, branchCheckDict["19"]));
            }

            // remove the first UNION in string
            String temmp = query.ToString().Remove(0, 7);
            _dt_grid_Trans.ExecuteQuery(temmp);
            // if no results from query delete first row which is empty
            if (_grid_Trans.DataTable.GetValue("DocEntry", 0).ToString() == "0")
                _grid_Trans.DataTable.Rows.Remove(0);
            _grid_Trans.Columns.Item("Checked").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
            _grid_Trans.Columns.Item("DocEntry").Editable = false;
            _grid_Trans.Columns.Item("DocNum").Editable = false;
            _grid_Trans.Columns.Item("ObjType").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
            SAPbouiCOM.ComboBoxColumn oCBC = (SAPbouiCOM.ComboBoxColumn)_grid_Trans.Columns.Item("ObjType");

            oCBC.ValidValues.Add("13", "A/R Invoice");
            oCBC.ValidValues.Add("14", "A/R Credit Note");
            oCBC.ValidValues.Add("18", "A/P Invoice");
            oCBC.ValidValues.Add("19", "A/P Credit Note");
            oCBC.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;

            _grid_Trans.Columns.Item("ObjType").Editable = false;
            _grid_Trans.Columns.Item("DocDate").Editable = false;
            _grid_Trans.Columns.Item("CardCode").Editable = false;
            _grid_Trans.Columns.Item("CardName").Editable = false;
            _grid_Trans.Columns.Item("NumAtCard").Editable = false;
            _grid_Trans.Columns.Item("DocCur").Editable = false;
            _grid_Trans.Columns.Item("DocTotal").Editable = false;

            // tree collaspable
            _grid_Trans.CollapseLevel = 1;
        }

        [FormEvent("ResizeAfter",false)]
        public static void OnAfterFormResize(SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = eCommon.SBO_Application.Forms.Item(pVal.FormUID);
            if (oForm.Items.Count > 0)
            {
                oForm.Items.Item("Item_1").Width = oForm.ClientWidth - 10;
                oForm.Items.Item("Item_1").Height = oForm.Items.Item("1").Top - 10 - oForm.Items.Item("Item_1").Top;
            }
        }

        [FormEvent("CloseBefore",true)]
        public static void OnBeforeFormClose(SAPbouiCOM.SBOItemEventArg pVal, out bool Bubble)
        {
            try
            {
                SAPbouiCOM.Form form = eCommon.SBO_Application.Forms.Item(pVal.FormUID);
                for (int i = 0; i < form.DataSources.DataTables.Count; i++)
                {
                    form.DataSources.DataTables.Item(i).Clear();
                }
            }
            finally
            {
                if (SBOAddon.oOpenForms.Contains(pVal.FormUID))
                    SBOAddon.oOpenForms.Remove(pVal.FormUID);

                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            Bubble = true;
        }

    }

   
}
