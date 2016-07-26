using System;
using System.Collections.Generic;
using System.Windows.Forms;
using SAPbouiCOM;
using SAPbobsCOM;

namespace TWM_KDS_AddOn
{
    class SBOAddon
    {
        public const string gcAddOnName = "TWM_KDS_AddOn";
        public const string gcAddonString = "TWM KDS AddOn";
        public static String WorkingDirectory = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\" + gcAddOnName;
        public static System.Collections.Hashtable Forms = new System.Collections.Hashtable();
        public static System.Collections.Specialized.OrderedDictionary oOpenForms = new System.Collections.Specialized.OrderedDictionary();
        public static System.Collections.Specialized.OrderedDictionary oFormEvents = new System.Collections.Specialized.OrderedDictionary();
        public static System.Collections.Hashtable oRegisteredFormEvents = new System.Collections.Hashtable();
        public Boolean Connected = true;
        public static String ParentUID = "";
        public static System.Resources.ResourceManager QueriesRM = new System.Resources.ResourceManager(String.Format("{0}.Src.Resource.Queries", System.Reflection.Assembly.GetExecutingAssembly().GetName().Name), System.Reflection.Assembly.GetExecutingAssembly());
        public enum BinStatus
        {
            NonStatus,
            Full,
            Partial,
            Empty
        }
        public enum Warehouses
        {
            Main,
            Lorry,
            Shop
        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(String[] Args)
        {
            try
            {
                SAPbouiCOM.SboGuiApi oGUI = new SAPbouiCOM.SboGuiApi();
                if (Args.Length == 0)
                {
                    oGUI.Connect("0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056");
                }
                else
                    oGUI.Connect(Args[0]);

                eCommon.SetApplication(oGUI.GetApplication(-1), gcAddOnName, true);

               // To add additional DB
                SBOAddon_DB addOnDb = new SBOAddon_DB();
                SBOAddon oAddOn = new SBOAddon();

                String str_directory = Environment.CurrentDirectory.ToString();
                String parent = System.IO.Directory.GetParent(System.IO.Directory.GetParent(System.IO.Directory.GetParent(str_directory).FullName).FullName).FullName+"\\";

               
                if (oAddOn.Connected)
                    System.Windows.Forms.Application.Run();
            }
            catch (Exception Ex)
            {
                System.Windows.Forms.MessageBox.Show("ERROR - Connection failed: " + Ex.Message); ;
            }

        }

        /// <summary>
        /// Constructor
        /// </summary>
        public SBOAddon()
        {
            try
            {
                //Application forms
                Forms = eCommon.CollectFormsAttribute();

                //--------------- remove and load menus -----------
                // Change if needed ---------------------------
                if (eCommon.SBO_Application.Menus.Exists(SBOAddon.gcAddOnName)) eCommon.SBO_Application.Menus.RemoveEx(SBOAddon.gcAddOnName);

                eCommon.SBO_Application.Menus.Item("43520").SubMenus.Add(SBOAddon.gcAddOnName, gcAddonString, BoMenuType.mt_POPUP, 99);
                // Change if needed ---------------------------

                foreach (string Key in Forms.Keys)
                {
                    FormAttribute oAttr = (FormAttribute)Forms[Key];
                    if (oAttr.HasMenu)
                    {
                       if (eCommon.SBO_Application.Menus.Exists(oAttr.FormType)) eCommon.SBO_Application.Menus.RemoveEx(oAttr.FormType);
                       eCommon.SBO_Application.Menus.Item(oAttr.ParentMenu).SubMenus.Add(oAttr.FormType, oAttr.MenuName, BoMenuType.mt_STRING, oAttr.Position);

                       //SAPbouiCOM.MenuCreationParams oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(eCommon.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
                       //oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                       //oCreationPackage.UniqueID = oAttr.FormType;
                       //oCreationPackage.String = oAttr.MenuName;
                       //oCreationPackage.Position = oAttr.Position;
                       //eCommon.SBO_Application.Menus.Item(oAttr.ParentMenu).SubMenus.AddEx(oCreationPackage);
                    }
                }

                try
                {
                    eCommon.SBO_Application.Menus.Item(SBOAddon.gcAddOnName).Image = Environment.CurrentDirectory + "\\Logo.JPG";
                }
                catch { }

                //Register Events
                RegisterAppEvents();
                RegisterFormEvents();

                //Register currently opened forms - initialized opened forms so it is ready to use.
                RegisterForms();

                //Need to change
                //Add SP
                //eCommon.AddSP("TWM_OPS_DBTrans_GET_DOCUMENT_LIST", SBOAddon.QueriesRM.GetString("TWM_OPS_DBTrans_GET_DOCUMENT_LIST"));

                //Need to change


                //Create Authorization
                AddAuthorizationTree();

                //Notify the users the addon is ready to use.
                eCommon.SBO_Application.StatusBar.SetText("Addon " + SBOAddon.gcAddOnName + " is ready.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

            }
            catch (Exception ex)
            {
                Connected = false;
                MessageBox.Show("Failed initializing addon. " + ex.Message);
            }
            finally
            {
            }

        }

        public void RegisterAppEvents()
        {
            eCommon.SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(OnMenuEvent);
            eCommon.SBO_Application.AppEvent += new _IApplicationEvents_AppEventEventHandler(OnAppEvents);
            eCommon.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
            
        }

        
        /// <summary>
        /// This method is only to keep a child form a modal!
        /// Could not find a way to make a form modal in the new SBO 9 UI.
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {

                switch (pVal.EventType)
                {
                    case BoEventTypes.et_ITEM_PRESSED:
                    case BoEventTypes.et_FORM_ACTIVATE:
                    case BoEventTypes.et_CLICK:
                        SAPbouiCOM.Form oForm = eCommon.SBO_Application.Forms.Item(FormUID);
                        //Check if a child form is opened for this form. if yes, dont click
                        //Find if a finder for this form exists... if it is then close it first
                        try
                        {
                            if (oForm.DataSources.UserDataSources.Item("ChildUID").Value != "")
                            {
                                SAPbouiCOM.Form oChildForm = eCommon.SBO_Application.Forms.Item(oForm.DataSources.UserDataSources.Item("ChildUID").Value);
                                oChildForm.Select();
                                BubbleEvent = false;
                                return;
                            }
                        }
                        catch { }

                        break;
                }
            }
            catch { }
        }


        /// <summary>
        /// Get the form events based on the Form Event attribute declared on the methods in each of the class
        /// </summary>
        public void RegisterFormEvents()
        {
            string NameSpace = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            Type FormEventAttrType = Type.GetType(string.Format("{0}.FormEventAttribute", NameSpace));

            foreach (System.Reflection.Assembly asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                if (asm.FullName.StartsWith("mscorlib"))
                    continue;
                if (asm.FullName.StartsWith("Interop"))
                    continue;
                if (asm.FullName.StartsWith("System"))
                    continue;
                if (asm.FullName.StartsWith("Microsoft"))
                    continue;

                foreach (Type type in asm.GetTypes())
                {

                    Type FormAttr = Type.GetType(string.Format("{0}.FormAttribute", NameSpace));
                    FormAttribute frmAttr = null;
                    foreach (System.Attribute Attr in type.GetCustomAttributes(FormAttr, false))
                    {
                        frmAttr = (FormAttribute)Attr;
                    }
                    //Get the methods attribute
                    foreach (System.Reflection.MethodInfo method in type.GetMethods())
                    {
                        foreach (System.Attribute Attr in method.GetCustomAttributes(FormEventAttrType, false))
                        {
                            SAPbouiCOM.EventForm oEvent = null;
                            FormEventAttribute frmEventAttr = (FormEventAttribute)Attr;
                            String sKey = string.Format("{0}_{1}", frmAttr.FormType, frmEventAttr.oEventType.ToString());
                            if (!SBOAddon.oFormEvents.Contains(frmAttr.FormType))
                            {
                                oEvent = eCommon.SBO_Application.Forms.GetEventForm(frmAttr.FormType);
                                SBOAddon.oFormEvents.Add(frmAttr.FormType, oEvent);
                            }
                            else
                            {
                                oEvent = (SAPbouiCOM.EventForm)SBOAddon.oFormEvents[frmAttr.FormType];
                            }

                            if (SBOAddon.oRegisteredFormEvents.Contains(sKey))
                                throw new Exception(string.Format("The form event method type [{0}] can not be registered twice", sKey));
                            else
                                SBOAddon.oRegisteredFormEvents.Add(sKey, "");

                            Type EventClass = oEvent.GetType();
                            System.Reflection.EventInfo oInfo = EventClass.GetEvent(frmEventAttr.oEventType.ToString());
                            if (oInfo == null)
                            {
                                throw new Exception(string.Format("Invalid method info name. [{0}]", frmEventAttr.oEventType.ToString()));
                            }
                            Delegate d = Delegate.CreateDelegate(oInfo.EventHandlerType, method);

                            oInfo.AddEventHandler(oEvent, d);

                        }

                    }
                }
            }

        }


        private void RegisterForms()
        {
            for (int i = 0; i < eCommon.SBO_Application.Forms.Count; i++)
            {
                if (!oOpenForms.Contains(eCommon.SBO_Application.Forms.Item(i).UniqueID))
                {
                    FormAttribute oAttrib = Forms[eCommon.SBO_Application.Forms.Item(i).TypeEx] as FormAttribute;
                    if (oAttrib != null)
                    {
                        try
                        {
                            //Execute the constructor
                            System.Reflection.Assembly asm = System.Reflection.Assembly.GetExecutingAssembly();
                            Type oType = asm.GetType(oAttrib.TypeName);
                            System.Reflection.ConstructorInfo ctor = oType.GetConstructor(new Type[1] { typeof(String) });
                            if (ctor != null)
                            {
                                object oForm = ctor.Invoke(new Object[1] { eCommon.SBO_Application.Forms.Item(i).UniqueID });
                            }
                            else
                                throw new Exception("No constructor which accepts the formUID found for form type - " + oAttrib.FormType);
                        }
                        catch (Exception ex)
                        {
                            eCommon.SBO_Application.MessageBox(ex.Message);
                        }
                    }
                }
            }
        }

        public void OnAppEvents(BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case BoAppEventTypes.aet_CompanyChanged:
                    if (eCommon.SBO_Application.Menus.Exists(SBOAddon.gcAddOnName)) eCommon.SBO_Application.Menus.RemoveEx(SBOAddon.gcAddOnName);
                    System.Windows.Forms.Application.Exit();
                    break;
                case BoAppEventTypes.aet_FontChanged:
                    break;
                case BoAppEventTypes.aet_LanguageChanged:
                    break;
                case BoAppEventTypes.aet_ServerTerminition:
                    break;
                case BoAppEventTypes.aet_ShutDown:
                    if (eCommon.SBO_Application.Menus.Exists(SBOAddon.gcAddOnName)) eCommon.SBO_Application.Menus.RemoveEx(SBOAddon.gcAddOnName);
                    System.Windows.Forms.Application.Exit();
                    break;
            }


        }

        public void OnMenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool Bubble)
        {
            Bubble = true;
            try
            {
                if (pVal.BeforeAction == true)
                {
                    SAPbouiCOM.Form oForm = null;

                    oForm = eCommon.SBO_Application.Forms.ActiveForm;
                    String sXML = oForm.GetAsXML();
                    switch (pVal.MenuUID)
                    {
                        case "1293":
                            break;
                        case "1285":        //Restore Form
                            break;
                    }
                }
                else
                {
                    //After Menu
                    SAPbouiCOM.Form oActiveForm = null;
                    switch (pVal.MenuUID)
                    {
                        case "1293":        //Delete Row Menu
                            oActiveForm = eCommon.SBO_Application.Forms.ActiveForm;
                            break;
                        case "1282":    //Add Menu pressed
                            oActiveForm = eCommon.SBO_Application.Forms.ActiveForm;
                            break;
                        case "1281":   //Find Menu
                            oActiveForm = eCommon.SBO_Application.Forms.ActiveForm;
                            break;
                        default:
                            FormAttribute oAttrib = Forms[pVal.MenuUID] as FormAttribute;
                            if (oAttrib != null)
                            {
                                try
                                {
                                    //Execute the constructor
                                    System.Reflection.Assembly asm = System.Reflection.Assembly.GetExecutingAssembly();
                                    Type oType = asm.GetType(oAttrib.TypeName);
                                    System.Reflection.ConstructorInfo ctor = oType.GetConstructor(new Type[0]);
                                    if (ctor != null)
                                    {
                                        object oForm = ctor.Invoke(new Object[0]);
                                    }
                                    else
                                        throw new Exception("No default constructor found for form type - " + oAttrib.FormType);
                                }
                                catch (Exception Ex)
                                {
                                    eCommon.SBO_Application.MessageBox(Ex.Message);
                                }
                            }
                            break;
                    }
                }
            }
            catch (Exception Ex)
            {
                eCommon.SBO_Application.StatusBar.SetText(Ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        private void AddAuthorizationTree()
        {
            try
            {
                SAPbobsCOM.UserPermissionTree oUserPer = (SAPbobsCOM.UserPermissionTree)eCommon.oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);
                int lErr = 0;

                if (oUserPer.GetByKey(SBOAddon.gcAddOnName) == false)
                {
                    oUserPer.PermissionID = SBOAddon.gcAddOnName;
                    oUserPer.Name = string.Format("Addon : {0}", SBOAddon.gcAddOnName);
                    oUserPer.Options = BoUPTOptions.bou_FullReadNone;
                    lErr = oUserPer.Add();
                    if (lErr != 0)
                        throw new Exception(eCommon.oCompany.GetLastErrorDescription());


                    System.Collections.Hashtable AuthTable = eCommon.CollectAuthorizationAttribute();
                    foreach (string FormType in AuthTable.Keys)
                    {
                        AuthorizationAttribute AuthAttrib = AuthTable[FormType] as AuthorizationAttribute;
                        oUserPer = (SAPbobsCOM.UserPermissionTree)eCommon.oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);
                        oUserPer.PermissionID = AuthAttrib.FormType;
                        oUserPer.Name = AuthAttrib.Name;
                        oUserPer.ParentID = AuthAttrib.ParentID;
                        oUserPer.Options = AuthAttrib.Options;
                        oUserPer.UserPermissionForms.FormType = AuthAttrib.FormType;
                        lErr = oUserPer.Add();
                        if (lErr != 0)
                            throw new Exception(eCommon.oCompany.GetLastErrorDescription());

                    }
                }


            }
            catch (Exception ex)
            {
                eCommon.oEventLog.WriteLine(DateTime.Now + eCommon.Filler + "Unable to create Authorization for " + SBOAddon.gcAddOnName + " Module. " + ex.Message);
            }
        }
     }
}