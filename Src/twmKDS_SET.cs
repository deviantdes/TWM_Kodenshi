using System;
using System.Collections.Generic;
using System.Text;


namespace TWM_KDS_AddOn
{
    [Form("twmKDS_SET", true, "Settings", SBOAddon.gcAddOnName,0)]
    //[Authorization("twmOPS_DBTST", "DB Trans Setup", SBOAddon.gcAddOnName, SAPbobsCOM.BoUPTOptions.bou_FullReadNone)]
    public class twmKDS_SET
    {
        SAPbouiCOM.Form _oForm = null;
        SAPbouiCOM.EditText _txtPath = null;
        SAPbouiCOM.CheckBox _cbPO = null;
        SAPbouiCOM.CheckBox _cbSO = null;
        SAPbouiCOM.Button _btnPath = null;
        SAPbouiCOM.Button _btnSave = null;

        public twmKDS_SET()
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

        public twmKDS_SET(SAPbouiCOM.Form oForm)
        {
            _oForm = oForm;
            GetItemReferences();
            if (!SBOAddon.oOpenForms.Contains(_oForm.UniqueID))
                SBOAddon.oOpenForms.Add(_oForm.UniqueID, this);
        }

        public twmKDS_SET(String FormUID)
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
                _txtPath.Value = SBOAddon_DB.Settings_xml_Path;
                _cbPO.Checked = SBOAddon_DB.Settings_Save_PO_Draft;
                _cbSO.Checked = SBOAddon_DB.Settings_Save_SO_Draft;

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
                _txtPath = _oForm.Items.Item("txtPath").Specific as SAPbouiCOM.EditText;
                _cbPO = _oForm.Items.Item("cbPO").Specific as SAPbouiCOM.CheckBox;
                _cbSO = _oForm.Items.Item("cbSO").Specific as SAPbouiCOM.CheckBox;
                _btnPath = _oForm.Items.Item("btnPath").Specific as SAPbouiCOM.Button;
                _btnSave = _oForm.Items.Item("btnSave").Specific as SAPbouiCOM.Button;

                _btnPath.PressedAfter += _btnPath_PressedAfter;
                _btnSave.PressedAfter += _btnSave_PressedAfter;
            }
            catch (Exception Ex){eCommon.SBO_Application.MessageBox(Ex.Message);}
        }

        void _btnSave_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (_txtPath.Value.Length < 1)
            {
                eCommon.SBO_Application.StatusBar.SetText("Invalid Path to save.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }
            else
            {
                Dictionary<String, String> dictionary = new Dictionary<String, String>();
                dictionary.Add("Export_XML_Path", _txtPath.Value);
                dictionary.Add("Save_PO_As_Draft", (_cbPO.Checked == true ? "1" : "0"));
                dictionary.Add("Save_SO_As_Draft", (_cbSO.Checked == true ? "1" : "0"));
              
                SBOAddon_DB.updateSettings(dictionary);
            }
        }

        void _btnPath_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SBOCustom.MyFileDialog fileDiag = new SBOCustom.MyFileDialog(eCommon.SBO_Application);
            _txtPath.Value = fileDiag.OpenFolderDialog() + "\\";
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
