using System;
using System.Collections.Generic;
using System.Text;


namespace TWM_KDS_AddOn
{
    [Form("twmKDS_UPTCNO", true, "Update Carton No.", SBOAddon.gcAddOnName,4)]
    //[Authorization("twmOPS_DBTST", "DB Trans Setup", SBOAddon.gcAddOnName, SAPbobsCOM.BoUPTOptions.bou_FullReadNone)]
    public class twmKDS_UPTCNO
    {
        SAPbouiCOM.Form _oForm = null;
        SAPbouiCOM.EditText _txtDocNum = null;
        SAPbouiCOM.Button _btnView = null;
        SAPbouiCOM.Button _btnUpdate = null;
        SAPbouiCOM.Grid _grid_Inv = null;
        SAPbouiCOM.DataTable _dt_grid_Inv = null;
        SAPbouiCOM.EditText _txtDocEntry = null;

        public twmKDS_UPTCNO()
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
                    catch(Exception ex) {
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

        public twmKDS_UPTCNO(SAPbouiCOM.Form oForm)
        {
            _oForm = oForm;
            GetItemReferences();
            if (!SBOAddon.oOpenForms.Contains(_oForm.UniqueID))
                SBOAddon.oOpenForms.Add(_oForm.UniqueID, this);
        }

        public twmKDS_UPTCNO(String FormUID)
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
                _grid_Inv.DataTable = _dt_grid_Inv;
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
                _txtDocNum = _oForm.Items.Item("txtDocNum").Specific as SAPbouiCOM.EditText;
                _txtDocEntry = _oForm.Items.Item("txtDocEn").Specific as SAPbouiCOM.EditText;
                _btnView = _oForm.Items.Item("btnView").Specific as SAPbouiCOM.Button;
                _btnUpdate = _oForm.Items.Item("btnUpdate").Specific as SAPbouiCOM.Button;
                _grid_Inv = _oForm.Items.Item("grid_Inv").Specific as SAPbouiCOM.Grid;
                _dt_grid_Inv = _oForm.DataSources.DataTables.Item("dt_grid_Inv");

                _txtDocNum.ChooseFromListAfter += _txtDocNum_ChooseFromListAfter;
                _btnView.PressedAfter += _btnView_PressedAfter;
                _btnUpdate.PressedAfter += _btnUpdate_PressedAfter;
            }
            catch (Exception Ex){eCommon.SBO_Application.MessageBox(Ex.Message);}
        }

        void _txtDocNum_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg pCFL = pVal as SAPbouiCOM.ISBOChooseFromListEventArg;

                if (pCFL.SelectedObjects != null)
                {
                    String DocNum = pCFL.SelectedObjects.GetValue("DocNum", 0).ToString();
                    String DocEntry = pCFL.SelectedObjects.GetValue("DocEntry", 0).ToString();
                    _oForm.DataSources.UserDataSources.Item("txtDocNum").ValueEx = DocNum;
                    _oForm.DataSources.UserDataSources.Item("txtDocEn").ValueEx = DocEntry;
                }
            }
            catch (Exception Ex) { eCommon.SBO_Application.StatusBar.SetText(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); }
        }

        // Update button
        void _btnUpdate_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (_grid_Inv.DataTable.GetValue("Doc No.", 0).ToString()=="0")
                    eCommon.SBO_Application.StatusBar.SetText("No Documents Updated.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                else
                {
                    updateDB();
                    eCommon.SBO_Application.StatusBar.SetText("Documents Updated.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }
            catch (Exception Ex) { eCommon.SBO_Application.StatusBar.SetText("No Documents to be Updated", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); }
        }


        //private void UpdateDB2()
        //{
        //    try
        //    {
        //        SAPbobsCOM.Recordset oRS = eCommon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
        //        oRS.DoQuery("SELECT TOP 1000 DocEntry, LineNum, U_TWM_CTBOXNO FROM INV1");
        //        StringBuilder strCase = new StringBuilder();
        //        StringBuilder strIn = new StringBuilder();

        //        int totalRows = _grid_Inv.Rows.Count;

        //        DateTime Start = DateTime.Now;
        //        System.Diagnostics.Debug.WriteLine("Start 1 : " + Start.ToString("HH:mm:ss fff"));
        //        for (int i = 0; i < oRS.RecordCount; i++)
        //        {
        //            strCase.Append(String.Format(" WHEN  DocEntry = '{0}' AND LineNum = '{1}' THEN '{2}'", oRS.Fields.Item("DocEntry").Value, oRS.Fields.Item("LineNum").Value, oRS.Fields.Item("U_TWM_CTBOXNO").Value));
        //            strIn.Append("'" + oRS.Fields.Item("DocEntry").Value + "',");

        //            oRS.MoveNext();
        //        }
        //        strIn.Append("0");
        //        String strQuery = String.Format(TWM_KDS_AddOn.Src.Resource.Queries.TWM_UPDATE_DB2, strCase.ToString(), strIn);

        //        SAPbobsCOM.Recordset oRS2 = eCommon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
        //        oRS2.DoQuery(strQuery);
        //        DateTime End = DateTime.Now;
        //        System.Diagnostics.Debug.WriteLine("End 1 : " + End.ToString("HH:mm:ss fff"));
        //        System.Diagnostics.Debug.WriteLine("Time 1: " + (End.Subtract(Start).TotalMilliseconds));

        //        oRS.MoveFirst();
        //        Start = DateTime.Now;
        //        System.Diagnostics.Debug.WriteLine("Start 2 : " + Start.ToString("HH:mm:ss fff"));
        //        for (int i = 0; i < oRS.RecordCount; i++)
        //        {
        //            String sSQL = String.Format("UPDATE INV1 SET U_TWM_CTBOXNO = '{2}' WHERE DocEntry = {0} AND LineNUm = {1}", oRS.Fields.Item("DocEntry").Value, oRS.Fields.Item("LineNum").Value, oRS.Fields.Item("U_TWM_CTBOXNO").Value);
        //            oRS2.DoQuery(sSQL);
        //            oRS.MoveNext();
        //        }
        //        End = DateTime.Now;
        //        System.Diagnostics.Debug.WriteLine("End 2 : " + End.ToString("HH:mm:ss fff"));
        //        System.Diagnostics.Debug.WriteLine("Time 2: " + (End.Subtract(Start).TotalMilliseconds));
        //        System.Diagnostics.Debug.WriteLine("Finish");
        //    }
        //    catch (Exception Ex)
        //    { }
        //}

        // Update carton box no.

        private void updateDB()
        {
            _btnUpdate.Item.Enabled = false;
            try
            {
                SAPbobsCOM.Recordset oRS = eCommon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                StringBuilder strCase = new StringBuilder();
                StringBuilder strIn = new StringBuilder();

                int totalRows = _grid_Inv.Rows.Count;

                for (int i = 0; i < totalRows; i++)
                {
                    strCase.Append(String.Format(" WHEN  DocEntry = '{0}' AND LineNum = '{1}' THEN '{2}'", _grid_Inv.DataTable.GetValue("DocEntry", i), _grid_Inv.DataTable.GetValue("LineNum", i), _grid_Inv.DataTable.GetValue("Carton Box No.", i)));
                    strIn.Append("'" + _oForm.DataSources.UserDataSources.Item("txtDocEn").ValueEx.ToString() + "',");
                }
                String strQuery = String.Format(TWM_KDS_AddOn.Src.Resource.Queries.TWM_UPDATE_DB2, strCase.ToString(), strIn.ToString().Remove(strIn.Length - 1));
                oRS.DoQuery(strQuery);
            }
            catch (Exception Ex) { eCommon.SBO_Application.StatusBar.SetText("Invalid Carton Box No. Value", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); }
            _btnUpdate.Item.Enabled = true;
        }
        
        // View button
        void _btnView_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            int temp = 0;
            if (_txtDocNum.Value.Length < 1)
                eCommon.SBO_Application.StatusBar.SetText("Enter a Document Number.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            else if(!Int32.TryParse(_txtDocNum.Value, out temp))
                eCommon.SBO_Application.StatusBar.SetText("Not a Valid Document Number", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            else
                populateGrid();
        }

        private void populateGrid()
        {
            try
            {
                _grid_Inv = _oForm.Items.Item("grid_Inv").Specific as SAPbouiCOM.Grid;
                String strQuery = String.Format(TWM_KDS_AddOn.Src.Resource.Queries.TWM_GET_INV, _oForm.DataSources.UserDataSources.Item("txtDocEn").ValueEx.ToString());
                _dt_grid_Inv.ExecuteQuery(strQuery);
                // if no results from query delete first row which is empty
                if (_grid_Inv.DataTable.GetValue("Doc No.", 0).ToString().Length < 1) _grid_Inv.DataTable.Rows.Remove(0);
                _grid_Inv.Columns.Item("DocEntry").Visible = false;
                _grid_Inv.Columns.Item("LineNum").Visible = false;
                _grid_Inv.Columns.Item("Doc No.").Editable = false;
                _grid_Inv.Columns.Item("Item Code").Editable = false;
                _grid_Inv.Columns.Item("BP Catalog No.").Editable = false;
                _grid_Inv.Columns.Item("Description").Editable = false;
                _grid_Inv.Columns.Item("Cust.PO No.").Editable = false;
                _grid_Inv.Columns.Item("Part No.").Editable = false;
                _grid_Inv.AutoResizeColumns();
                // link button for itemNo
                SAPbouiCOM.EditTextColumn oEditColItemNo;
                oEditColItemNo = ((SAPbouiCOM.EditTextColumn)(_grid_Inv.Columns.Item("Item Code")));
                oEditColItemNo.LinkedObjectType = "4";


            }
            catch (Exception Ex) { eCommon.SBO_Application.StatusBar.SetText(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); }
        }

        [FormEvent("ResizeAfter",false)]
        public static void OnAfterFormResize(SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = eCommon.SBO_Application.Forms.Item(pVal.FormUID);
            if (oForm.Items.Count > 0)
            {
                
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
