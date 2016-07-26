using System;
using System.Collections.Generic;
using System.Text;

namespace TWM_KDS_AddOn
{
    using SAPbobsCOM;
    using SAPbouiCOM;
    using B1WizardBase;
    using System;
    public class SBOAddon_DB:B1Db
    {
        private static Boolean _settings_Save_PO_Draft = true;
        private static Boolean _settings_Save_SO_Draft = true;
        private static String _settings_xml_Path = "C:\\";
        private static SAPbobsCOM.Recordset ors;

        public SBOAddon_DB():base() 
        {
            ors = eCommon.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            try
            {
                // try see if table exist
                ors.DoQuery("SELECT * FROM [@TWM_SETTINGS]");
                while(!ors.EoF)
                {
                    if (ors.Fields.Item("U_TWM_Settings_Type").Value.ToString() == "Export_XML_Path")
                        _settings_xml_Path = ors.Fields.Item("U_TWM_Settings_Value").Value.ToString();
                    else if (ors.Fields.Item("U_TWM_Settings_Type").Value.ToString() == "Save_PO_As_Draft")
                        _settings_Save_PO_Draft = (ors.Fields.Item("U_TWM_Settings_Value").Value.ToString() == "1");
                    else if (ors.Fields.Item("U_TWM_Settings_Type").Value.ToString() == "Save_SO_As_Draft")
                        _settings_Save_SO_Draft = (ors.Fields.Item("U_TWM_Settings_Value").Value.ToString() == "1");

                    ors.MoveNext();
                }
            }
            catch 
            {
                //One or more metadata not found. try to recreate them.
                if (ors != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ors);
                    ors = null;
                    GC.Collect();
                }

                Tables = new B1DbTable[] { 
                               new B1DbTable("@TWM_SETTINGS", "Global Settings table", BoUTBTableType.bott_NoObject)
                };

                Columns = new B1DbColumn[] 
                       {    
                           new B1DbColumn("@TWM_SETTINGS","TWM_Settings_Type","Setting Type",BoFieldTypes.db_Alpha,BoFldSubTypes.st_None,250,new B1WizardBase.B1DbValidValue[0], -1),
                           new B1DbColumn("@TWM_SETTINGS","TWM_Settings_Value","Setting Value",BoFieldTypes.db_Alpha,BoFldSubTypes.st_None,250,new B1WizardBase.B1DbValidValue[0], -1),
                           new B1DbColumn("@TWM_SETTINGS","TWM_Settings_AddOn","Setting for which AddOn",BoFieldTypes.db_Alpha,BoFldSubTypes.st_None,250,new B1WizardBase.B1DbValidValue[0],-1)
                       };

                try
                {
                    eCommon.SBO_Application.MetadataAutoRefresh = false;
                    this.Add(eCommon.oCompany);
                    addDefaultSettings();
                }
                catch (Exception ex){eCommon.SBO_Application.MessageBox(ex.Message);}
                finally
                {
                    eCommon.SBO_Application.MetadataAutoRefresh = true;
                }
            }

            if (ors != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ors);
                ors = null;
                GC.Collect();
            }
        }

        private void addDefaultSettings()
        {
            try
            {
                ors = eCommon.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                ors.DoQuery(String.Format("INSERT INTO [@TWM_SETTINGS] ([Code],[Name],[U_TWM_Settings_Type],[U_TWM_Settings_Value],[U_TWM_Settings_AddOn]) VALUES ('1','1','Export_XML_Path','{0}','TWM_KDS_AddOn')", "C:\\"));
                ors.DoQuery(String.Format("INSERT INTO [@TWM_SETTINGS] ([Code],[Name],[U_TWM_Settings_Type],[U_TWM_Settings_Value],[U_TWM_Settings_AddOn]) VALUES ('2','2','Save_PO_As_Draft','{0}','TWM_KDS_AddOn')", "1"));
                ors.DoQuery(String.Format("INSERT INTO [@TWM_SETTINGS] ([Code],[Name],[U_TWM_Settings_Type],[U_TWM_Settings_Value],[U_TWM_Settings_AddOn]) VALUES ('3','3','Save_SO_As_Draft','{0}','TWM_KDS_AddOn')", "1"));
            }
            catch
            {
                throw;
            }
        }

        public static void updateSettings(Dictionary<String, String> dict_Settings)
        {
            try
            {
                ors = eCommon.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                foreach (KeyValuePair<string, string> entry in dict_Settings)
                {
                    ors.DoQuery(String.Format("UPDATE [@TWM_SETTINGS] SET U_TWM_Settings_Value='{0}' WHERE U_TWM_Settings_Type='{1}';", entry.Value,entry.Key));

                    if (entry.Key == "Export_XML_Path")
                        _settings_xml_Path = entry.Value;
                    else if (entry.Key == "Save_PO_As_Draft")
                        _settings_Save_PO_Draft = (entry.Value=="1");
                    else if (entry.Key == "Save_SO_As_Draft")
                        _settings_Save_SO_Draft = (entry.Value=="1");
                }
                eCommon.SBO_Application.StatusBar.SetText("Settings Updated !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch { }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ors);
                ors = null;
                GC.Collect();
            }
        }

        public static Boolean Settings_Save_PO_Draft
        {
            get { return SBOAddon_DB._settings_Save_PO_Draft; }
            private set { SBOAddon_DB._settings_Save_PO_Draft = value; }
        }

        public static Boolean Settings_Save_SO_Draft
        {
            get { return SBOAddon_DB._settings_Save_SO_Draft; }
            private set { SBOAddon_DB._settings_Save_SO_Draft = value; }
        }

        public static String Settings_xml_Path
        {
            get { return SBOAddon_DB._settings_xml_Path; }
            private set { SBOAddon_DB._settings_xml_Path = value; }
        }
    }
}
