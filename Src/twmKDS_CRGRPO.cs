using System;
using System.Collections.Generic;
using System.Text;


namespace TWM_KDS_AddOn
{
    [Form("twmKDS_CRGRPO", true, "Multiple PO to GPRO", SBOAddon.gcAddOnName, 1)]
    //[Authorization("twmOPS_DBTST", "DB Trans Setup", SBOAddon.gcAddOnName, SAPbobsCOM.BoUPTOptions.bou_FullReadNone)]
    public class twmKDS_CRGRPO
    {
        SAPbouiCOM.Form _oForm = null;

        SAPbouiCOM.Button _btnView = null;
        SAPbouiCOM.Button _btnNext = null;
        SAPbouiCOM.Button _btnBack = null;
        SAPbouiCOM.Button _btnPost = null;
        //Screen 1 texts
        SAPbouiCOM.EditText _txtVendor = null;
        SAPbouiCOM.EditText _txtDTFRM = null;
        SAPbouiCOM.EditText _txtDTTO = null;
        SAPbouiCOM.EditText _txtSrchPO = null;
        SAPbouiCOM.EditText _txtTTLQTY = null;
        SAPbouiCOM.EditText _txtTTLAMT = null;
        SAPbouiCOM.EditText _txtNAME = null;
        SAPbouiCOM.EditText _txtPSTDT = null;
        SAPbouiCOM.EditText _txtDUEDT = null;
        SAPbouiCOM.EditText _txtDOCDT = null;
        SAPbouiCOM.EditText _txtRMK = null;
        SAPbouiCOM.DataTable _dt_VendorMatrix = null;
        SAPbouiCOM.DataTable _dt_VendorGrid = null;
        SAPbouiCOM.CheckBox _cbDraft = null;
        SAPbouiCOM.Matrix _mat_Vendor = null;
        SAPbouiCOM.Grid _grid_Vendor = null;

        private bool toggleGridCheckBox = false;

        public twmKDS_CRGRPO()
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

        public twmKDS_CRGRPO(SAPbouiCOM.Form oForm)
        {
            _oForm = oForm;
            GetItemReferences();
            if (!SBOAddon.oOpenForms.Contains(_oForm.UniqueID))
                SBOAddon.oOpenForms.Add(_oForm.UniqueID, this);
        }

        public twmKDS_CRGRPO(String FormUID)
        {
            _oForm = eCommon.SBO_Application.Forms.Item(FormUID);
            GetItemReferences();
            if (!SBOAddon.oOpenForms.Contains(_oForm.UniqueID))
                SBOAddon.oOpenForms.Add(_oForm.UniqueID, this);

        }

        private void InitForm()
        {
            _oForm.Freeze(true);
            try
            {
                _oForm.DataSources.UserDataSources.Item("txtPSTDT").ValueEx = DateTime.Now.ToString("yyyyMMdd");
                //_oForm.DataSources.UserDataSources.Item("txtDUEDT").ValueEx = DateTime.Now.ToString("yyyyMMdd");
                _oForm.DataSources.UserDataSources.Item("txtDOCDT").ValueEx = DateTime.Now.ToString("yyyyMMdd");

                _grid_Vendor.DataTable = _dt_VendorGrid;

                SAPbouiCOM.ChooseFromList oCFL = _oForm.ChooseFromLists.Item("cflVendor");
                SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
                SAPbouiCOM.Condition oCon = oCons.Add();
                oCon.Alias = "CardType";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "S";

                oCFL.SetConditions(oCons);               
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
                _txtVendor = _oForm.Items.Item("txtVendor").Specific as SAPbouiCOM.EditText;
                _txtDTFRM = _oForm.Items.Item("txtDTFRM").Specific as SAPbouiCOM.EditText;
                _txtDTTO = _oForm.Items.Item("txtDTTO").Specific as SAPbouiCOM.EditText;
                _btnView = _oForm.Items.Item("btnView").Specific as SAPbouiCOM.Button;
                _txtSrchPO = _oForm.Items.Item("txtSrchPO").Specific as SAPbouiCOM.EditText;
                _txtTTLAMT = _oForm.Items.Item("txtTtlAmt").Specific as SAPbouiCOM.EditText;
                _txtTTLQTY = _oForm.Items.Item("txtTtlQty").Specific as SAPbouiCOM.EditText;

                _txtNAME = _oForm.Items.Item("txtNAME").Specific as SAPbouiCOM.EditText;
                _txtPSTDT = _oForm.Items.Item("txtPSTDT").Specific as SAPbouiCOM.EditText;
                _txtDUEDT = _oForm.Items.Item("txtDUEDT").Specific as SAPbouiCOM.EditText;
                _txtDOCDT = _oForm.Items.Item("txtDOCDT").Specific as SAPbouiCOM.EditText;
                _txtRMK = _oForm.Items.Item("txtRMK").Specific as SAPbouiCOM.EditText;

                _btnPost = _oForm.Items.Item("btnPost").Specific as SAPbouiCOM.Button;
                _btnNext = _oForm.Items.Item("btnNext").Specific as SAPbouiCOM.Button;
                _btnBack = _oForm.Items.Item("btnBack").Specific as SAPbouiCOM.Button;
                _cbDraft = _oForm.Items.Item("cbDraft").Specific as SAPbouiCOM.CheckBox;

                _mat_Vendor = _oForm.Items.Item("mat_Vendor").Specific as SAPbouiCOM.Matrix;
                
                _mat_Vendor.AutoResizeColumns();
                _dt_VendorMatrix = _oForm.DataSources.DataTables.Item("dt_VendorfromOPOR");

                _grid_Vendor = _oForm.Items.Item("grid_Ven").Specific as SAPbouiCOM.Grid;
                _dt_VendorGrid = _oForm.DataSources.DataTables.Item("dtListLines");

                _mat_Vendor.LinkPressedBefore += _mat_Vendor_LinkPressedBefore;
                _txtSrchPO.KeyDownAfter += _txtSrchPO_KeyDownAfter;
                _txtVendor.ChooseFromListAfter += _txtVendor_ChooseFromListAfter;
                _btnView.PressedBefore += _btnView_PressedBefore;
                _btnNext.PressedAfter += _btnNext_PressedAfter;
                _btnBack.PressedAfter += _btnBack_PressedAfter;
                _grid_Vendor.PressedAfter += _grid_Vendor_PressedAfter;
                _grid_Vendor.ValidateAfter += _grid_Vendor_ValidateAfter;
                _btnPost.PressedAfter += _btnPost_PressedAfter;
                // Double click events
                _grid_Vendor.DoubleClickAfter += _grid_Vendor_DoubleClickAfter;
                _mat_Vendor.DoubleClickAfter += _mat_Vendor_DoubleClickAfter;
            }
            catch (Exception Ex){eCommon.SBO_Application.MessageBox(Ex.Message);}
        }

        void _mat_Vendor_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (pVal.Row == 0 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_CTRL)
                return;

            if (pVal.ColUID == "colCHK" && pVal.Row == 0)
                checkAllMat(_mat_Vendor);
        }

        // toggle all checkbox in matrix
        private void checkAllMat(SAPbouiCOM.Matrix theMatrix)
        {
            try
            {
                _oForm.Freeze(true);
                if (!toggleGridCheckBox)
                {
                    for (int i = 0; i < _dt_VendorMatrix.Rows.Count; i++)
                    {
                        _dt_VendorMatrix.SetValue("Check", i, "Y");
                    }

                    theMatrix.LoadFromDataSource();

                    toggleGridCheckBox = true;
                }
                else
                {
                    for (int i = 0; i < _dt_VendorMatrix.Rows.Count; i++)
                    {
                        _dt_VendorMatrix.SetValue("Check", i, "N");
                    }

                    theMatrix.LoadFromDataSource();

                    toggleGridCheckBox = false;
                }
            }
            catch (Exception ex) { eCommon.SBO_Application.MessageBox(ex.Message); }

            _oForm.Freeze(false);
        }

        // double click to select all checkbox in grid
        void _grid_Vendor_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (pVal.Row == -1 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_CTRL)
                return;

            if (pVal.ColUID == "Checked" && pVal.Row == -1)
                checkAllGrid(_grid_Vendor);       
        }

        // toggle all checkbox in grid
        private void checkAllGrid(SAPbouiCOM.Grid theGrid)
        {
            _oForm.Freeze(true);

            if (!toggleGridCheckBox)
            {
                for (int i = 0; i < theGrid.Rows.Count; i++)
                {
                    theGrid.DataTable.SetValue("Checked",i,"Y");
                }

                toggleGridCheckBox = true;
            }
            else
            {
                for (int i = 0; i < theGrid.Rows.Count; i++){
                    theGrid.DataTable.SetValue("Checked", i, "N");
                }
                toggleGridCheckBox = false;
            }

            _oForm.Freeze(false);
        }

        // TODO: Submit the PO to create GRPO 
        void _btnPost_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (eCommon.DataTableIndexOf(_dt_VendorGrid, "Checked", "Y") == null)
                    eCommon.SBO_Application.StatusBar.SetText("No PO Selected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                else
                {
                    SAPbobsCOM.Documents oReceipt = null;
                    // If save as draft
                    if (_cbDraft.Checked)
                    {
                        oReceipt = eCommon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts) as SAPbobsCOM.Documents;
                        oReceipt.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes;
                    }
                    // If save as GRPO
                    else
                        oReceipt = eCommon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes) as SAPbobsCOM.Documents;


                    oReceipt.CardCode = _oForm.DataSources.UserDataSources.Item("txtVendor").ValueEx;
                    oReceipt.CardName = _oForm.DataSources.UserDataSources.Item("txtNAME").ValueEx;

                    oReceipt.DocDate = DateTime.ParseExact(_oForm.DataSources.UserDataSources.Item("txtPSTDT").ValueEx, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.CurrentInfo, System.Globalization.DateTimeStyles.AssumeLocal);
                    //check if empty
                    if (_oForm.DataSources.UserDataSources.Item("txtDOCDT").ValueEx.Length > 0)
                        oReceipt.DocDueDate = DateTime.ParseExact(_oForm.DataSources.UserDataSources.Item("txtDOCDT").ValueEx, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.CurrentInfo, System.Globalization.DateTimeStyles.AssumeLocal);
                    //check if empty
                    if (_oForm.DataSources.UserDataSources.Item("txtDUEDT").ValueEx.Length > 0)
                        oReceipt.TaxDate = DateTime.ParseExact(_oForm.DataSources.UserDataSources.Item("txtDUEDT").ValueEx, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.CurrentInfo, System.Globalization.DateTimeStyles.AssumeLocal);
                    //check if empty
                    if (_oForm.DataSources.UserDataSources.Item("txtRMK").ValueEx.Length > 0)
                        oReceipt.Comments = _oForm.DataSources.UserDataSources.Item("txtRMK").ValueEx;

                    int[] SelectedRows = eCommon.DataTableIndexOf(_dt_VendorGrid, "Checked", "Y");
                    int iLineCount = 0;

                    if (SelectedRows != null)
                        foreach (int Row in SelectedRows)
                        {
                            //Check whether it is checked. if not just continue;
                            //first time theres already a new line, so only add on 2nd line on wards
                            if (iLineCount > 0) oReceipt.Lines.Add();
                            // the document type
                            String temp = _dt_VendorGrid.GetValue("Doc_Type", Row).ToString();
                            oReceipt.Lines.BaseType = Int32.Parse(_dt_VendorGrid.GetValue("Doc_Type", Row).ToString());
                            // the document id
                            oReceipt.Lines.BaseEntry = Int32.Parse(_dt_VendorGrid.GetValue("Doc_Entry", Row).ToString());
                            // which line is it in the document
                            oReceipt.Lines.BaseLine = Int32.Parse(_dt_VendorGrid.GetValue("Line_Num", Row).ToString());
                            oReceipt.Lines.Quantity = double.Parse(_dt_VendorGrid.GetValue("Quantity", Row).ToString());
                            oReceipt.Lines.UserFields.Fields.Item("U_TWM_CTBOXNO").Value = _dt_VendorGrid.GetValue("Checked", Row);

                            iLineCount++;

                        }

                    //add the invoice
                    int iErr = oReceipt.Add();
                    //open the invoice
                    if (iErr != 0)
                    {
                        throw new Exception(eCommon.oCompany.GetLastErrorDescription());
                    }
                    else
                    {
                        String iDocEntry = eCommon.oCompany.GetNewObjectKey();
                        String sObjectType = eCommon.oCompany.GetNewObjectType();

                        oReceipt = eCommon.oCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)int.Parse(sObjectType)) as SAPbobsCOM.Documents;
                        oReceipt.GetByKey(int.Parse(iDocEntry));
                        String sDocNum = oReceipt.DocNum.ToString();

                        eCommon.SBO_Application.OpenForm((SAPbouiCOM.BoFormObjectEnum)int.Parse(sObjectType), "", iDocEntry);
                        _oForm.Close();

                    }
                }
            }
            catch (Exception Ex) { eCommon.SBO_Application.StatusBar.SetText(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); }
           
        }

        // update fields after qty or price has been change
        void _grid_Vendor_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            switch (pVal.ColUID)
            {
                case "Quantity":case "Unit Price":case "Total Carton Box":
                    _grid_Vendor.DataTable.Columns.Item("Total Amount").Cells.Item(pVal.Row).Value = (((double)_grid_Vendor.DataTable.Columns.Item("Quantity").Cells.Item(pVal.Row).Value) * ((double)_grid_Vendor.DataTable.Columns.Item("Unit Price").Cells.Item(pVal.Row).Value));
                    updateFields();
                    break;
            }
        }

        void _grid_Vendor_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (pVal.ColUID == "Checked")
                updateFields();
        }

        //update fields
        private void updateFields()
        {
            int[] SelectedRows = eCommon.DataTableIndexOf(_dt_VendorGrid, "Checked", "Y");
            double dAmount = 0;
            double dQty = 0;

            if (SelectedRows != null)
                foreach (int Row in SelectedRows)
                {
                    dAmount += (Double)_dt_VendorGrid.GetValue("Total Amount", Row);
                    dQty += (Double)_dt_VendorGrid.GetValue("Quantity", Row);
                }

            _txtTTLAMT.Value = dAmount.ToString();
            _txtTTLQTY.Value = dQty.ToString();
        }

        void _btnBack_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            changePane(1);
        }

        // on button next check which rows are checked, and proceed to pane 2
        void _btnNext_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                _mat_Vendor.FlushToDataSource();
                int[] SelectedRows = eCommon.DataTableIndexOf(_dt_VendorMatrix, "Check", "Y");

                if (SelectedRows == null || SelectedRows.Length < 1)
                    eCommon.SBO_Application.StatusBar.SetText("No PO selected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                else
                {
                    changePane(2);
                }
            }
            catch (Exception Ex) { eCommon.SBO_Application.StatusBar.SetText(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); }  
        }

        private void changePane(int Paneno)
        {
            toggleGridCheckBox = false;
            switch (Paneno)
            {
                case 1:
                    _oForm.PaneLevel = Paneno;
                    _oForm.ActiveItem = "txtSrchPO";
                    _oForm.Items.Item("txtVendor").Enabled = true;
                    break;
                case 2:
                    populateGrid("N");
                    _txtTTLAMT.Value = "";
                    _txtTTLQTY.Value = "";
                    _oForm.PaneLevel = Paneno;
                    _oForm.ActiveItem = "txtNAME";
                    _oForm.Items.Item("txtVendor").Enabled = false;
                    _cbDraft.Checked = SBOAddon_DB.Settings_Save_PO_Draft;
                    break;
                default:
                    break;
            }

        }

        // Search and highlight rows
        void _txtSrchPO_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (_txtSrchPO.Value.Length > 0)
            try
            {
                _oForm.Freeze(true);
                int[] selectedRows = twmKDS_CRINV.DataTableIndexOf(_dt_VendorMatrix, "NUMATCARD", _txtSrchPO.String);

                //to select multiple rows on search
                _mat_Vendor.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                if (selectedRows != null)
                {
                    for (int i = 0; i < selectedRows.Length; i++)
                        _mat_Vendor.SelectRow(selectedRows[i] + 1, true, true);
                    eCommon.SBO_Application.StatusBar.SetText(selectedRows.Length + " match(es) for search", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                else
                    eCommon.SBO_Application.StatusBar.SetText("No Match", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                _oForm.Freeze(false);
            }
            catch (Exception Ex) { eCommon.SBO_Application.StatusBar.SetText(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); }
        }

        // link button in matrix, SONO column
        void _mat_Vendor_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.ColUID == "colSONO")
            {
                String docEntry = (_mat_Vendor.GetCellSpecific("colDOCET", pVal.Row) as SAPbouiCOM.EditText).String;
                eCommon.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_PurchaseOrder, "", docEntry);
                BubbleEvent = false;
            }
        }

        //view button and then populate matrix
        void _btnView_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (_txtVendor.String.Length < 1 || _txtDTFRM.String.Length < 1 || _txtDTTO.String.Length < 1)
                {
                    eCommon.SBO_Application.StatusBar.SetText("Do not leave blanks for the Selections", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    BubbleEvent = false;
                }
                else
                    populateMatrix();
            }
            catch (Exception Ex)
            {
                BubbleEvent = false;
                eCommon.SBO_Application.StatusBar.SetText(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        // query and populate grid AND link button
        private void populateGrid(String check)
        {
            try
            {
                int[] SelectedRows = eCommon.DataTableIndexOf(_dt_VendorMatrix, "Check", "Y");

                StringBuilder oSB = new StringBuilder();
                foreach (int Row in SelectedRows)
                {
                    oSB.Append(String.Format("{0},", _dt_VendorMatrix.GetValue("DocEntry", Row)));
                }
                String strQuery = String.Format(TWM_KDS_AddOn.Src.Resource.Queries.TWM_GET_PO_GRID, oSB.Append("0").ToString(), check);
                _dt_VendorGrid.ExecuteQuery(strQuery);
                // if no results from query delete first row which is empty
                if (_grid_Vendor.DataTable.GetValue("ItemNo", 0).ToString().Length < 1) _grid_Vendor.DataTable.Rows.Remove(0);
                _grid_Vendor.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                _grid_Vendor.Columns.Item("ItemNo").Editable = false;
                _grid_Vendor.Columns.Item("Item Description").Editable = false;
                _grid_Vendor.Columns.Item("Total Amount").Editable = false;
                _grid_Vendor.Columns.Item("Tax Code").Editable = false;
                _grid_Vendor.Columns.Item("PO Customer").Editable = false;
                _grid_Vendor.Columns.Item("Warehouse").Editable = false;
                _grid_Vendor.Columns.Item("DeliveryDate").Editable = false;
                _grid_Vendor.Columns.Item("Doc_Type").Visible = false;
                _grid_Vendor.Columns.Item("Doc_Entry").Visible = false;
                _grid_Vendor.Columns.Item("Line_Num").Visible = false;

                // link button for itemNo
                SAPbouiCOM.EditTextColumn oEditColItemNo;
                oEditColItemNo = ((SAPbouiCOM.EditTextColumn)(_grid_Vendor.Columns.Item("ItemNo")));
                oEditColItemNo.LinkedObjectType = "4";

                // link button for WareHouse
                SAPbouiCOM.EditTextColumn oEditColWarehouse;
                oEditColWarehouse = ((SAPbouiCOM.EditTextColumn)(_grid_Vendor.Columns.Item("Warehouse")));
                oEditColWarehouse.LinkedObjectType = "64";


            }
            catch (Exception Ex) { eCommon.SBO_Application.StatusBar.SetText(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); }
        }

        // query and populate matrix
        private void populateMatrix()
        {
            try
            {
                String strQuery = String.Format(TWM_KDS_AddOn.Src.Resource.Queries.TWM_GET_PO_MATRIX, _txtVendor.String, _txtDTFRM.Value, _txtDTTO.Value);

                 _dt_VendorMatrix = _oForm.DataSources.DataTables.Item("dt_VendorfromOPOR");
                 _dt_VendorMatrix.ExecuteQuery(strQuery);

                
                 if (_dt_VendorMatrix.Rows.Count > 0)
                 {
                     _mat_Vendor.Columns.Item("col_0").DataBind.Bind("dt_VendorfromOPOR", "Row");
                     _mat_Vendor.Columns.Item("colVCC").DataBind.Bind("dt_VendorfromOPOR", "CardCode");
                     _mat_Vendor.Columns.Item("colVCN").DataBind.Bind("dt_VendorfromOPOR", "CardName");
                     _mat_Vendor.Columns.Item("colCHK").DataBind.Bind("dt_VendorfromOPOR", "Check");
                     _mat_Vendor.Columns.Item("colSONO").DataBind.Bind("dt_VendorfromOPOR", "DocNum");
                     _mat_Vendor.Columns.Item("colPONO").DataBind.Bind("dt_VendorfromOPOR", "NumAtCard");
                     _mat_Vendor.Columns.Item("colDOCET").DataBind.Bind("dt_VendorfromOPOR", "DocEntry");
                     _mat_Vendor.LoadFromDataSource();
                     _mat_Vendor.AutoResizeColumns();
                 }
                
                
            }
            catch (Exception ex) { eCommon.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); }
        }

        void _txtVendor_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg pCFL = pVal as SAPbouiCOM.ISBOChooseFromListEventArg;
                if (pCFL.SelectedObjects != null)
                {
                    String CardCode = pCFL.SelectedObjects.GetValue("CardCode", 0).ToString();
                    _oForm.DataSources.UserDataSources.Item("txtVendor").ValueEx = CardCode;
                    // For screen 2's name field
                    _oForm.DataSources.UserDataSources.Item("txtNAME").ValueEx = pCFL.SelectedObjects.GetValue("CardName", 0).ToString();
                }
            }
            catch (Exception Ex) { eCommon.SBO_Application.StatusBar.SetText(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); }
   
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
