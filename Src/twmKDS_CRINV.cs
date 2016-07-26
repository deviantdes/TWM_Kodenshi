using System;
using System.Collections.Generic;
using System.Text;


namespace TWM_KDS_AddOn
{
    [Form("twmKDS_CRINV", true, "Multiple SO to Invoice", SBOAddon.gcAddOnName, 1)]
    //[Authorization("twmOPS_DBTST", "DB Trans Setup", SBOAddon.gcAddOnName, SAPbobsCOM.BoUPTOptions.bou_FullReadNone)]
    public class twmKDS_CRINV
    {
        SAPbouiCOM.Form _oForm = null;

        SAPbouiCOM.Button _btnView = null;
        SAPbouiCOM.Button _btnNext = null;
        SAPbouiCOM.Button _btnPost = null;
        SAPbouiCOM.Button _btnBack = null;

        //Screen 1 texts
        SAPbouiCOM.EditText _txtCSTCD = null;
        SAPbouiCOM.EditText _txtDTFRM = null;
        SAPbouiCOM.EditText _txtDTTO = null;
        SAPbouiCOM.EditText _txtSrchPO = null;

        //Screen 2
        SAPbouiCOM.EditText _txtNAME = null;
        SAPbouiCOM.EditText _txtPSTDT = null;
        SAPbouiCOM.EditText _txtDUEDT = null;
        SAPbouiCOM.EditText _txtDOCDT = null;
        SAPbouiCOM.EditText _txtRMK = null;
        SAPbouiCOM.EditText _txtTTLQTY = null;
        SAPbouiCOM.EditText _txtTTLAMT = null;
        SAPbouiCOM.EditText _txtTTLCBOX = null;
        SAPbouiCOM.CheckBox _cbDraft = null;
        SAPbouiCOM.Matrix _mat_CST = null;
        SAPbouiCOM.Grid _grid_CST = null;
        SAPbouiCOM.DataTable _dt_CST = null;
        SAPbouiCOM.DataTable _dt_LST = null;
        private bool toggleGridCheckBox = false;

        public twmKDS_CRINV()
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

        public twmKDS_CRINV(SAPbouiCOM.Form oForm)
        {
            _oForm = oForm;
            GetItemReferences();
            if (!SBOAddon.oOpenForms.Contains(_oForm.UniqueID))
                SBOAddon.oOpenForms.Add(_oForm.UniqueID, this);
        }

        public twmKDS_CRINV(String FormUID)
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

                _mat_CST.Columns.Item("colDOCET").Visible = false;

                _grid_CST.DataTable = _dt_LST;

                SAPbouiCOM.ChooseFromList oCFL = _oForm.ChooseFromLists.Item("cflCSTCD");
                SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
                SAPbouiCOM.Condition oCon = oCons.Add();
                oCon.Alias = "CardType";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "C";

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
                _btnView = _oForm.Items.Item("btnView").Specific as SAPbouiCOM.Button;
                _txtCSTCD = _oForm.Items.Item("txtCSTCD").Specific as SAPbouiCOM.EditText;
                _txtDTFRM = _oForm.Items.Item("txtDTFRM").Specific as SAPbouiCOM.EditText;
                _txtDTTO = _oForm.Items.Item("txtDTTO").Specific as SAPbouiCOM.EditText;
                _txtSrchPO = _oForm.Items.Item("txtSrchPO").Specific as SAPbouiCOM.EditText;
                _btnNext = _oForm.Items.Item("btnNext").Specific as SAPbouiCOM.Button;
                _btnBack = _oForm.Items.Item("btnBack").Specific as SAPbouiCOM.Button;
                _txtTTLAMT = _oForm.Items.Item("txtTtlAmt").Specific as SAPbouiCOM.EditText;
                _txtTTLQTY = _oForm.Items.Item("txtTtlQty").Specific as SAPbouiCOM.EditText;
                _txtTTLCBOX = _oForm.Items.Item("txtTtlCrt").Specific as SAPbouiCOM.EditText;

                _txtNAME = _oForm.Items.Item("txtNAME").Specific as SAPbouiCOM.EditText;
                _txtPSTDT = _oForm.Items.Item("txtPSTDT").Specific as SAPbouiCOM.EditText;
                _txtDUEDT = _oForm.Items.Item("txtDUEDT").Specific as SAPbouiCOM.EditText;
                _txtDOCDT = _oForm.Items.Item("txtDOCDT").Specific as SAPbouiCOM.EditText;
                _txtRMK = _oForm.Items.Item("txtRMK").Specific as SAPbouiCOM.EditText;
                _cbDraft = _oForm.Items.Item("cbDraft").Specific as SAPbouiCOM.CheckBox;
                _btnPost = _oForm.Items.Item("btnPost").Specific as SAPbouiCOM.Button;
                _mat_CST = _oForm.Items.Item("mat_CST").Specific as SAPbouiCOM.Matrix;
                // Make the matrix fill the area
                _mat_CST.AutoResizeColumns();

                //initialize the data table first
                _dt_CST = _oForm.DataSources.DataTables.Item("dt_SelectCusfromORDR");
                _dt_LST = _oForm.DataSources.DataTables.Item("dtListLines");

                _grid_CST = _oForm.Items.Item("grid_CST").Specific as SAPbouiCOM.Grid;

                _btnPost.PressedAfter += _btnPost_PressedAfter;
                _btnBack.PressedAfter += _btnBack_PressedAfter;
                _btnNext.PressedAfter += _btnNext_PressedAfter;
                _btnView.PressedBefore += _btnView_PressedBefore;
                _txtCSTCD.ChooseFromListAfter += _txtCSTCD_ChooseFromListAfter;
                _txtSrchPO.KeyDownAfter += _txtSrchPO_KeyDownAfter;
                _grid_CST.ValidateAfter += _grid_CST_ValidateAfter;
                _grid_CST.PressedAfter += _grid_CST_PressedAfter;
                _mat_CST.LinkPressedBefore += _mat_CST_LinkPressedBefore;
                // Double click events
                _grid_CST.DoubleClickAfter += _grid_CST_DoubleClickAfter;
                _mat_CST.DoubleClickAfter += _mat_CST_DoubleClickAfter;
            }
            catch (Exception Ex){eCommon.SBO_Application.MessageBox(Ex.Message);}
        }

        void _mat_CST_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (pVal.Row == 0 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_CTRL)
                return;

            if (pVal.ColUID == "colCHK" && pVal.Row == 0)
                checkAllMat(_mat_CST);
        }

        // toggle all checkbox in matrix
        private void checkAllMat(SAPbouiCOM.Matrix theMatrix)
        {
            try
            {
                _oForm.Freeze(true);
                if (!toggleGridCheckBox)
                {
                    for (int i = 0; i < _dt_CST.Rows.Count; i++)
                    {
                        _dt_CST.SetValue("Check", i, "Y");
                    }
                    _mat_CST.LoadFromDataSource();

                    toggleGridCheckBox = true;
                }
                else
                {
                
                    for (int i = 0; i < _dt_CST.Rows.Count; i++)
                    {
                        _dt_CST.SetValue("Check", i, "N");
                    }

                    _mat_CST.LoadFromDataSource();

                    toggleGridCheckBox = false;
                }
            }
            catch (Exception ex){ eCommon.SBO_Application.MessageBox(ex.Message);}

            _oForm.Freeze(false);
        }

        // double click to select all checkbox in grid
        void _grid_CST_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (pVal.Row == -1 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_CTRL)
                return;

            if (pVal.ColUID == "Checked" && pVal.Row == -1)
                checkAllGrid(_grid_CST);
        }

        // toggle all checkbox in grid
        private void checkAllGrid(SAPbouiCOM.Grid theGrid)
        {
            _oForm.Freeze(true);

            if (!toggleGridCheckBox)
            {
                for (int i = 0; i < theGrid.Rows.Count; i++)
                {
                    theGrid.DataTable.SetValue("Checked", i, "Y");
                }

                toggleGridCheckBox = true;
            }
            else
            {
                for (int i = 0; i < theGrid.Rows.Count; i++)
                {
                    theGrid.DataTable.SetValue("Checked", i, "N");
                }
                toggleGridCheckBox = false;
            }

            _oForm.Freeze(false);
        }

        //open link buttun in matrix using DocEntry
        void _mat_CST_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.ColUID == "colSONO")
            {
                String docEntry = (_mat_CST.GetCellSpecific("colDOCET", pVal.Row) as SAPbouiCOM.EditText).String;
                eCommon.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_Order, "", docEntry);
                BubbleEvent = false;
            }

        }

        // To pass the SO to invoice
        void _btnPost_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (eCommon.DataTableIndexOf(_dt_LST, "Checked", "Y") == null)
                    eCommon.SBO_Application.StatusBar.SetText("No SO Selected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                else
                {

                    SAPbobsCOM.Documents oInv = null;
                    // If save as draft
                    if (_cbDraft.Checked)
                    {
                        oInv = eCommon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts) as SAPbobsCOM.Documents;
                        oInv.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices;
                    }
                    // If save as Invoice
                    else
                        oInv = eCommon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices) as SAPbobsCOM.Documents;

                    oInv.CardCode = _oForm.DataSources.UserDataSources.Item("txtCSTCD").ValueEx;
                    oInv.CardName = _oForm.DataSources.UserDataSources.Item("txtNAME").ValueEx;
               
                    oInv.DocDate = DateTime.ParseExact(_oForm.DataSources.UserDataSources.Item("txtPSTDT").ValueEx, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.CurrentInfo, System.Globalization.DateTimeStyles.AssumeLocal);
                    //check if empty
                    if (_oForm.DataSources.UserDataSources.Item("txtDOCDT").ValueEx.Length>0)
                    oInv.DocDueDate = DateTime.ParseExact(_oForm.DataSources.UserDataSources.Item("txtDOCDT").ValueEx, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.CurrentInfo, System.Globalization.DateTimeStyles.AssumeLocal);
                    //check if empty
                    if (_oForm.DataSources.UserDataSources.Item("txtDUEDT").ValueEx.Length > 0)
                    oInv.TaxDate = DateTime.ParseExact(_oForm.DataSources.UserDataSources.Item("txtDUEDT").ValueEx, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.CurrentInfo, System.Globalization.DateTimeStyles.AssumeLocal);
                    //check if empty
                    if (_oForm.DataSources.UserDataSources.Item("txtRMK").ValueEx.Length > 0)
                    oInv.Comments = _oForm.DataSources.UserDataSources.Item("txtRMK").ValueEx;

                    int[] SelectedRows = eCommon.DataTableIndexOf(_dt_LST, "Checked", "Y");
                    int iLineCount = 0;

                    if (SelectedRows != null)
                    foreach (int Row in SelectedRows)
                    {
                       //Check whether it is checked. if not just continue;
                       //first time theres already a new line, so only add on 2nd line on wards
                       if (iLineCount>0) oInv.Lines.Add();
                         // the document type
                         oInv.Lines.BaseType = Int32.Parse(_dt_LST.GetValue("Doc_Type", Row).ToString());
                         // the document id
                         oInv.Lines.BaseEntry = Int32.Parse(_dt_LST.GetValue("Doc_Entry", Row).ToString());
                         // which line is it in the document
                         oInv.Lines.BaseLine = Int32.Parse(_dt_LST.GetValue("Line_Num", Row).ToString());
                         oInv.Lines.Quantity = double.Parse(_dt_LST.GetValue("Quantity", Row).ToString());
                         oInv.Lines.UserFields.Fields.Item("U_TWM_CTBOXNO").Value = _dt_LST.GetValue("Checked", Row);

                         iLineCount++;

                    }

                    //add the invoice
                    int iErr = oInv.Add();
                    //open the invoice
                    if (iErr != 0)
                    {
                        throw new Exception(eCommon.oCompany.GetLastErrorDescription());
                    }
                    else
                    {
                        String iDocEntry = eCommon.oCompany.GetNewObjectKey();
                        String sObjectType = eCommon.oCompany.GetNewObjectType();

                        oInv = eCommon.oCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)int.Parse(sObjectType)) as SAPbobsCOM.Documents;
                        oInv.GetByKey(int.Parse(iDocEntry));
                        String sDocNum = oInv.DocNum.ToString();

                        eCommon.SBO_Application.OpenForm((SAPbouiCOM.BoFormObjectEnum)int.Parse(sObjectType), "", iDocEntry);
                        _oForm.Close();

                    }
                }
            }
            catch (Exception Ex) { eCommon.SBO_Application.StatusBar.SetText(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); }
           
        }

        void _grid_CST_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (pVal.ColUID == "Checked")
                updateFields();
        }

        // on update of quantity or unit price column, update the total amount
        void _grid_CST_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            switch (pVal.ColUID)
            {
                case "Quantity":case"Unit Price":case"Total Carton Box":
                         _grid_CST.DataTable.Columns.Item("Total Amount").Cells.Item(pVal.Row).Value = (((double)_grid_CST.DataTable.Columns.Item("Quantity").Cells.Item(pVal.Row).Value) * ((double)_grid_CST.DataTable.Columns.Item("Unit Price").Cells.Item(pVal.Row).Value));
                         updateFields(); 
                         break;
            }
        }

        //update fields
        private void updateFields()
        {
            int[] SelectedRows = eCommon.DataTableIndexOf(_dt_LST, "Checked", "Y");
            double dAmount = 0;
            double dQty = 0;
            int dCrtBox = 0;

            if (SelectedRows != null)
                foreach (int Row in SelectedRows)
                {
                    dAmount += (Double)_dt_LST.GetValue("Total Amount", Row);
                    dQty += (Double)_dt_LST.GetValue("Quantity", Row);
                    int temp;
                    if (int.TryParse(_dt_LST.GetValue("Total Carton Box", Row).ToString(), out temp))
                    {
                        dCrtBox += temp;
                    }
                    
                }

            _txtTTLAMT.Value = dAmount.ToString();
            _txtTTLQTY.Value = dQty.ToString();
            _txtTTLCBOX.Value = dCrtBox.ToString();
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
                //listofDocEntries = new List<String>();
                //for (int i = 1; i <= _mat_CST.RowCount; i++)
                //{
                //    SAPbouiCOM.CheckBox check = (SAPbouiCOM.CheckBox)_mat_CST.Columns.Item("colCHK").Cells.Item(i).Specific;
                //    SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)_mat_CST.Columns.Item("colDOCET").Cells.Item(i).Specific;   
                //    if (check.Checked)                
                //        listofDocEntries.Add(oEditText.Value);
                //}
                _mat_CST.FlushToDataSource();
                int[] SelectedRows = eCommon.DataTableIndexOf(_dt_CST, "Check", "Y");

                if (SelectedRows==null || SelectedRows.Length < 1) 
                    eCommon.SBO_Application.StatusBar.SetText("No SO selected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                else
                {
                    changePane(2);
                }
            }
            catch (Exception Ex){eCommon.SBO_Application.StatusBar.SetText(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);}
        }

        // Search and highlight rows
        void _txtSrchPO_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (_txtSrchPO.Value.Length > 0)
            try
            {
                _oForm.Freeze(true);
                 int[] selectedRows = DataTableIndexOf(_dt_CST, "NUMATCARD", _txtSrchPO.String);
                 _mat_CST.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                 if (selectedRows != null)
                 {
                     for (int i = 0; i < selectedRows.Length; i++)
                         _mat_CST.SelectRow(selectedRows[i] + 1, true, true);          
                     eCommon.SBO_Application.StatusBar.SetText(selectedRows.Length+" match(es) for search", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                 }
                 else
                     eCommon.SBO_Application.StatusBar.SetText("No Match", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                 _oForm.Freeze(false);
            }
            catch (Exception Ex) {eCommon.SBO_Application.StatusBar.SetText(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);}
        }

        // VIEW button click and populate matrix
        void _btnView_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (_txtCSTCD.String.Length < 1 || _txtDTFRM.String.Length < 1 || _txtDTTO.String.Length < 1)
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

        private void changePane(int Paneno)
        {
            toggleGridCheckBox = false;
            switch (Paneno)
            {
                case 1:
                    _oForm.PaneLevel = Paneno;
                    _oForm.ActiveItem = "txtSrchPO";
                    _oForm.Items.Item("txtCSTCD").Enabled = true;                    
                    break;
                case 2:
                    populateGrid("N");
                    _txtTTLAMT.Value = "";
                    _txtTTLQTY.Value = "";
                    _txtTTLCBOX.Value = "";
                    _oForm.PaneLevel = Paneno;
                    _oForm.ActiveItem = "txtNAME";
                    _oForm.Items.Item("txtCSTCD").Enabled = false;
                    _cbDraft.Checked = SBOAddon_DB.Settings_Save_SO_Draft;
                    break;
                default:
                    break;
            }
          
        }

        private void populateGrid(String check)
        {
            try
            {
                  int[] SelectedRows = eCommon.DataTableIndexOf(_dt_CST, "Check", "Y");

                  StringBuilder oSB = new StringBuilder();
                  foreach(int Row in SelectedRows)
                  {
                      oSB.Append(String.Format("{0},", _dt_CST.GetValue("DocEntry",Row)));   
                  }
                  String strQuery = String.Format(TWM_KDS_AddOn.Src.Resource.Queries.TWM_GET_SO_GRID, oSB.Append("0").ToString(), check);
                  _dt_LST.ExecuteQuery(strQuery);
                  // if no results from query delete first row which is empty
                  if (_grid_CST.DataTable.GetValue("ItemNo", 0).ToString().Length < 1) _grid_CST.DataTable.Rows.Remove(0);
                  _grid_CST.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                  _grid_CST.Columns.Item("ItemNo").Editable = false;
                  _grid_CST.Columns.Item("BP Catalog No").Editable = false;
                  _grid_CST.Columns.Item("Item Description").Editable = false;
                  _grid_CST.Columns.Item("Total Amount").Editable = false;
                  _grid_CST.Columns.Item("Warehouse").Editable = false;
                  _grid_CST.Columns.Item("Delivery Date").Editable = false;
                  _grid_CST.Columns.Item("Doc_Type").Visible = false;
                  _grid_CST.Columns.Item("Doc_Entry").Visible = false;
                  _grid_CST.Columns.Item("Line_Num").Visible = false;

                 
                  // link button for itemNo
                  SAPbouiCOM.EditTextColumn oEditColItemNo;
                  oEditColItemNo = ((SAPbouiCOM.EditTextColumn)(_grid_CST.Columns.Item("ItemNo")));
                  oEditColItemNo.LinkedObjectType = "4";

                  // link button for WareHouse
                  SAPbouiCOM.EditTextColumn oEditColWarehouse;
                  oEditColWarehouse = ((SAPbouiCOM.EditTextColumn)(_grid_CST.Columns.Item("Warehouse")));
                  oEditColWarehouse.LinkedObjectType = "64";

                  updateFields();
            }
            catch (Exception Ex){eCommon.SBO_Application.StatusBar.SetText(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);}
        }

        private void populateMatrix()
        {
            try{
                String strQuery = String.Format(TWM_KDS_AddOn.Src.Resource.Queries.TWM_GET_SO_MATRIX, _txtCSTCD.String, _txtDTFRM.Value, _txtDTTO.Value);

                _dt_CST = _oForm.DataSources.DataTables.Item("dt_SelectCusfromORDR");
                _dt_CST.ExecuteQuery(strQuery);

                if (_dt_CST.Rows.Count > 0)
                {
                    _mat_CST.Columns.Item("col_0").DataBind.Bind("dt_SelectCusfromORDR", "Row");
                    _mat_CST.Columns.Item("colCSTCD").DataBind.Bind("dt_SelectCusfromORDR", "CardCode");
                    _mat_CST.Columns.Item("colCSTNM").DataBind.Bind("dt_SelectCusfromORDR", "CardName");
                    _mat_CST.Columns.Item("colCHK").DataBind.Bind("dt_SelectCusfromORDR", "Check");
                    _mat_CST.Columns.Item("colSONO").DataBind.Bind("dt_SelectCusfromORDR", "DocNum");
                    _mat_CST.Columns.Item("colPONO").DataBind.Bind("dt_SelectCusfromORDR", "NumAtCard");
                    _mat_CST.Columns.Item("colDOCET").DataBind.Bind("dt_SelectCusfromORDR", "DocEntry");
                    _mat_CST.LoadFromDataSource();
                    _mat_CST.AutoResizeColumns();
                }
            }
            catch(Exception ex){eCommon.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);}
        }

        //Used for the search filter, this returns the rows that contains the search value
        public static int[] DataTableIndexOf(SAPbouiCOM.DataTable oDT, string ColumnUID, string SearchValue)
        {
     
            int[] iResult = null;
            string sDT = oDT.SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly).ToUpper();
            //Normalize the SearchValue first : \, [, ^, $, the period or dot .,  |,  ?,  *,  +,  (,  )
            string NormSearchValue = SearchValue.ToUpper().Replace("\\", "\\\\");
            NormSearchValue = NormSearchValue.Replace("[", "\\[");
            NormSearchValue = NormSearchValue.Replace("^", "\\^");
            NormSearchValue = NormSearchValue.Replace("$", "\\$");
            NormSearchValue = NormSearchValue.Replace(".", "\\.");
            NormSearchValue = NormSearchValue.Replace("|", "\\|");
            NormSearchValue = NormSearchValue.Replace("?", "\\?");
            NormSearchValue = NormSearchValue.Replace("*", "\\*");
            NormSearchValue = NormSearchValue.Replace("+", "\\+");
            NormSearchValue = NormSearchValue.Replace("(", "\\(");
            NormSearchValue = NormSearchValue.Replace(")", "\\)");


            string SearchString = string.Format("<Cell><ColumnUid>{0}</ColumnUid><Value>{1}".ToUpper(), ColumnUID.ToUpper(), NormSearchValue);
            System.Text.RegularExpressions.Regex oRegex = new System.Text.RegularExpressions.Regex(SearchString);
            System.Text.RegularExpressions.MatchCollection oMatches = oRegex.Matches(sDT);

            iResult = new int[oMatches.Count];
            for (int i = 0; i < oMatches.Count; i++)
            {
                System.Text.RegularExpressions.Match oMatch = oMatches[i];
                SearchString = "<ROW>";
                oRegex = new System.Text.RegularExpressions.Regex(SearchString);
                System.Text.RegularExpressions.MatchCollection oRowMatches = oRegex.Matches(sDT.Substring(0, oMatch.Index));

                iResult[i] = oRowMatches.Count - 1;
            }

            if (iResult.Length == 0)
                return null;
            else
                return iResult;
        }

        //Choose the value from the opened list
        private void _txtCSTCD_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg pCFL = pVal as SAPbouiCOM.ISBOChooseFromListEventArg;

                if (pCFL.SelectedObjects != null)
                {
                    String CardCode = pCFL.SelectedObjects.GetValue("CardCode", 0).ToString();
                    _oForm.DataSources.UserDataSources.Item("txtCSTCD").ValueEx = CardCode;
                    _oForm.DataSources.UserDataSources.Item("txtNAME").ValueEx = pCFL.SelectedObjects.GetValue("CardName", 0).ToString();
                }
            }
            catch (Exception Ex){eCommon.SBO_Application.StatusBar.SetText(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);}
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
