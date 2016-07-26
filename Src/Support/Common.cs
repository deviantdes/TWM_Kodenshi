namespace TWM_KDS_AddOn
{
    using SAPbobsCOM;
    using SAPbouiCOM;
    using B1WizardBase;
    using System;
    using System.Windows.Forms;
    using System.Management;
    using System.Runtime.InteropServices;
    
    public class eCommon
    {
        public static System.IO.StreamWriter oEventLog = null;
        private static SAPbouiCOM.Application _App = null;
        private static SAPbobsCOM.Company _oCom = null;
        public const string Filler = "      ";
        public static int iPriceAccuracy = 2;
        public static int iQtyAccuracy = 0;
        public static int iAmountAccuracy = 2;
        public static string LocalCurr = "";
        public static string SystemCurr = "";
        public static string TempField = "";
        public static bool isSegmentedAcc = false;
        public static string sAcctSeparator = "";

        public static SAPbouiCOM.Application SBO_Application
        {
            get
            {
                return _App;
            }
        }
        public static SAPbobsCOM.Company oCompany
        {
            get
            {
                return _oCom;
            }
        }

        /// <summary>
        /// this method is for init company wihtout UI connection
        /// </summary>
        /// <param name="ServerType"></param>
        /// <param name="Server"></param>
        /// <param name="LicServer"></param>
        /// <param name="CompanyDB"></param>
        /// <param name="UserName"></param>
        /// <param name="Password"></param>
        public static void InitCompany(SAPbobsCOM.BoDataServerTypes ServerType, String Server, String LicServer, String CompanyDB, String UserName, String Password)
        {
            //Create a Log File
            CreateLogFile();

            _oCom = new SAPbobsCOM.Company();
            _oCom.DbServerType = ServerType;
            _oCom.Server = Server;
            _oCom.LicenseServer = LicServer;
            _oCom.CompanyDB = CompanyDB;
            _oCom.UserName = UserName;
            _oCom.Password = Password;

            _oCom.Connect();
            if (!_oCom.Connected)
                throw new Exception(_oCom.GetLastErrorDescription());

        }


        #region "Create user form"

        public static void SetApplication(SAPbouiCOM.Application oApp, String AddOnName, bool Multiple)
        {
            _App = oApp;
            _oCom = new SAPbobsCOM.Company();
            string cookie = _oCom.GetContextCookie();
            string connInfo = _App.Company.GetConnectionContext(cookie);
            int retCode = _oCom.SetSboLoginContext(connInfo);
            if (retCode == 0)
            {
                if (Multiple)
                    _oCom = (SAPbobsCOM.Company)_App.Company.GetDICompany();
                else
                {
                    _oCom.Connect();
                }
                //Create a Log File
                CreateLogFile();

                // CHECK TWM ADDON LICENCE
#if !DEBUG
                if (!IsDesignTime())
                {
                    if (!CheckLicenceEX(AddOnName))
                    {
                        throw new System.Runtime.InteropServices.COMException("No Addon Licence.");
                    }
                }
#endif

                //Decimal Accuracies
                SAPbobsCOM.CompanyService oCompanyInfo = _oCom.GetCompanyService();
                iPriceAccuracy = oCompanyInfo.GetAdminInfo().PriceAccuracy;
                iQtyAccuracy = oCompanyInfo.GetAdminInfo().AccuracyofQuantities;
                iAmountAccuracy = oCompanyInfo.GetAdminInfo().TotalsAccuracy;


                //Initialize some values
                SBObob Sbob = (SBObob)eCommon.oCompany.GetBusinessObject(BoObjectTypes.BoBridge);
                SAPbobsCOM.Recordset oRS = Sbob.GetLocalCurrency();
                eCommon.LocalCurr = (string)oRS.Fields.Item(0).Value;
                oRS = Sbob.GetSystemCurrency();
                eCommon.SystemCurr = (string)oRS.Fields.Item(0).Value;
                oRS = (SAPbobsCOM.Recordset)eCommon.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRS.DoQuery("SELECT COUNT(*) FROM OASG");
                if ((int)oRS.Fields.Item(0).Value > 0)
                {
                    eCommon.isSegmentedAcc = true;
                    eCommon.sAcctSeparator = oCompanyInfo.GetAdminInfo().AccountSegmentsSeparator;
                }
                else
                {
                    eCommon.isSegmentedAcc = false;
                }

                eCommon.ReleaseComObject((object)oCompanyInfo);
                eCommon.ReleaseComObject((object)oRS);
                eCommon.ReleaseComObject((object)Sbob);


            }
            else
            {
                throw new Exception(_oCom.GetLastErrorDescription());
            }
        }


        private static void CreateLogFile()
        {
            try
            {
                //Logging part
                string sLogDir = Environment.GetFolderPath( Environment.SpecialFolder.CommonApplicationData) + "\\" +  SBOAddon.gcAddOnName +  "\\EventLog";
                if (!System.IO.Directory.Exists(sLogDir)) System.IO.Directory.CreateDirectory(sLogDir);
                string sLogFN = sLogDir + "\\EventLog" + System.DateTime.Now.ToString("yyyyMM") + ".txt";

                int iCount = 1;
                while (true)
                {
                    try
                    {
                        //oEventLog = System.IO.File.AppendText(sLogFN)
                        if (iCount == 1)
                            InitializeEventLog(sLogFN);
                        else
                        {
                            sLogFN = sLogFN.Replace(".txt", "");
                            InitializeEventLog(sLogFN + "(" + iCount + ").txt");
                        }
                        break;
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.ToLower().Contains("because it is being used by another process"))
                        {
                            iCount += 1;
                            if (iCount > 100)
                            {
                                eCommon.SBO_Application.StatusBar.SetText(SBOAddon.gcAddOnName + " Unable to start logging. Starting without logging.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                break;
                            }
                        }
                        else
                        {
                            eCommon.SBO_Application.StatusBar.SetText(SBOAddon.gcAddOnName + " Could not start logging." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                            break;
                        }
                    }
                }

                if (eCommon.oEventLog != null)
                {
                    eCommon.oEventLog.AutoFlush = true;

                    eCommon.oEventLog.WriteLine("________________________________________ NEW SESSION ________________________________________");
                    if (iCount > 1) eCommon.oEventLog.Write(DateTime.Now.ToString("yyyyMMdd HH:mm:ss") + "  " + iCount + " concurrent sessions detected");
                    eCommon.oEventLog.WriteLine(DateTime.Now.ToString("yyyyMMdd HH:mm:ss") + "  Start Event Log");
                    eCommon.oEventLog.WriteLine(DateTime.Now.ToString("yyyyMMdd HH:mm:ss") + "  Initiating Add On");
                }
            }
            catch { }

        }

        private static void InitializeEventLog(string sLogFN)
        {
            //If Not System.IO.File.Exists(sLogFN) Then System.IO.File.Create(sLogFN)
            eCommon.oEventLog = System.IO.File.AppendText(sLogFN);
        }

        public static string GetChildFormUID(string FormUID)
        {
            String sChildUID = "";
            try
            {
                sChildUID = SBO_Application.Forms.Item(FormUID).DataSources.UserDataSources.Item("ChldUID").Value;
            }
            catch
            { }

            if (sChildUID != "")
                sChildUID = GetChildFormUID(sChildUID);
            else
                return FormUID;

            return sChildUID;
        }

        public static double GetCurrencyRate(SAPbobsCOM.Company TheCompany, String currency, DateTime refDate)
        {
            SAPbobsCOM.SBObob vObj = TheCompany.GetBusinessObject(BoObjectTypes.BoBridge) as SAPbobsCOM.SBObob;
            SAPbobsCOM.Recordset rs = TheCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            rs = vObj.GetCurrencyRate(currency, refDate);
            Double dRate = (double)rs.Fields.Item(0).Value;
            if (dRate == 0)
            {
                String errResult = TheCompany.GetLastErrorDescription();
            }

            return dRate;

        }

        public static string GetDefaultSeries(SAPbobsCOM.BoObjectTypes ObjType)
        {
            SAPbobsCOM.CompanyService oCmpSrv = null;
            SAPbobsCOM.SeriesService oSeriesService = null;
            SAPbobsCOM.Series oSeries = null;
            SAPbobsCOM.DocumentTypeParams oDocumentTypeParams = null;

            try
            {
                //get company service
                oCmpSrv = oCompany.GetCompanyService();

                //get series service
                oSeriesService = (SAPbobsCOM.SeriesService)oCmpSrv.GetBusinessService(ServiceTypes.SeriesService);

                //get new series
                oSeries = oSeriesService.GetDataInterface(SeriesServiceDataInterfaces.ssdiSeries) as SAPbobsCOM.Series;

                //get DocumentTypeParams for filling the document type
                oDocumentTypeParams = oSeriesService.GetDataInterface(SeriesServiceDataInterfaces.ssdiDocumentTypeParams) as SAPbobsCOM.DocumentTypeParams;

                //set the document type (e.g. A/R Invoice=13)
                oDocumentTypeParams.Document = ((int)ObjType).ToString();

                //get the default series of the SaleOrder documentset the document type
                oSeries = oSeriesService.GetDefaultSeries(oDocumentTypeParams);

                return oSeries.Series.ToString();
            }
            catch
            {
                return "-1";
            }
            finally
            {
                eCommon.ReleaseComObject(oCmpSrv);
                eCommon.ReleaseComObject(oSeriesService);
                eCommon.ReleaseComObject(oSeries);
                eCommon.ReleaseComObject(oDocumentTypeParams);
            }
        }

        public static string GetDefaultSeries(String UDOType)
        {
            SAPbobsCOM.CompanyService oCmpSrv = null;
            SAPbobsCOM.SeriesService oSeriesService = null;
            SAPbobsCOM.Series oSeries = null;
            SAPbobsCOM.DocumentTypeParams oDocumentTypeParams = null;

            try
            {
                //get company service
                oCmpSrv = oCompany.GetCompanyService();

                //get series service
                oSeriesService = (SAPbobsCOM.SeriesService)oCmpSrv.GetBusinessService(ServiceTypes.SeriesService);

                //get new series
                oSeries = oSeriesService.GetDataInterface(SeriesServiceDataInterfaces.ssdiSeries) as SAPbobsCOM.Series;

                //get DocumentTypeParams for filling the document type
                oDocumentTypeParams = oSeriesService.GetDataInterface(SeriesServiceDataInterfaces.ssdiDocumentTypeParams) as SAPbobsCOM.DocumentTypeParams;

                //set the document type (e.g. A/R Invoice=13)
                oDocumentTypeParams.Document = UDOType;

                //get the default series of the SaleOrder documentset the document type
                oSeries = oSeriesService.GetDefaultSeries(oDocumentTypeParams);

                return oSeries.Series.ToString();
            }
            catch
            {
                return "-1";
            }
            finally
            {
                eCommon.ReleaseComObject(oCmpSrv);
                eCommon.ReleaseComObject(oSeriesService);
                eCommon.ReleaseComObject(oSeries);
                eCommon.ReleaseComObject(oDocumentTypeParams);
            }
        }

        /// <summary>
        /// Get the row index of any data inside an SAPbouiCOM.DataTable
        /// </summary>
        /// <param name="oDT">The DataTable to check</param>
        /// <param name="ColumnUID">The ColumnUID in the datatable to check the content for</param>
        /// <param name="SearchValue">Exact string to search in the column</param>
        /// <returns>Returns null if no match extists</returns>
        public static int[] DataTableIndexOf(SAPbouiCOM.DataTable oDT, string ColumnUID, string SearchValue)
        {
            int[] iResult = null;
            string sDT = oDT.SerializeAsXML(BoDataTableXmlSelect.dxs_DataOnly).ToUpper();
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


        public static string DrawForm(string FileName, SAPbouiCOM.Form ParentForm)
        {
            string FormUID = "";
            try
            {
                int Xloc = 0;
                int Yloc = 0;
                if (ParentForm != null)
                {
                    try
                    {
                        Xloc = ParentForm.Left + 50;
                        Yloc = ParentForm.Top + 20;
                    }
                    catch { }
                }
                FormUID = LoadFromXML(FileName, Xloc, Yloc);

                if (FormUID != "")
                {
                    SBO_Application.Forms.Item(FormUID).Freeze(true);
                    try
                    {
                        SBO_Application.Forms.Item(FormUID).Visible = true;
                        //AddDataSource(FormUID);
                        //BindDataSource(FormUID);
                        //InitializeForm(FormUID, ExtendedInfo);
                    }
                    catch
                    {
                        SBO_Application.Forms.Item(FormUID).Freeze(false);
                    }
                }

                return FormUID;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ":" + ex.StackTrace.ToString());
                return "";
            }
            finally
            {
                if (FormUID != "") SBO_Application.Forms.Item(FormUID).Freeze(false);
            }
        }



        public static string LoadFromXML(string FileName, int Xloc, int Yloc)
        {
            string sXml = "";
            if (System.IO.File.Exists(FileName))
            {
                System.Xml.XmlDocument oXmlDoc = new System.Xml.XmlDocument();

                // load the content of the XML File
                string sPath = System.Windows.Forms.Application.StartupPath.ToString();
                try
                {
                    oXmlDoc.Load(sPath + "\\" + FileName);
                }
                catch (Exception ex)
                {
                    if (eCommon.oEventLog != null) eCommon.oEventLog.WriteLine(DateTime.Now.ToString("yyyyMMdd HH:mm:ss") + Filler + "EX.MESSAGE : " + ex.Message);
                }
                sXml = oXmlDoc.InnerXml;
            }
            else
            {
                String ResourceName = string.Format("{0}.Src.Resource.{1}", System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, FileName);
                sXml = eCommon.GetXMLResource(ResourceName);
            }

            //Modify the form size to match the client Font Size
            sXml = ModifySize(sXml);

            //Modify Location of form.
            if (Xloc != 0)
                sXml = ModifyXLoc(sXml, Xloc);

            if (Yloc != 0)
                ModifyYLoc(sXml, Yloc);



            // load the form to the SBO application in one batch
            SBO_Application.LoadBatchActions(ref sXml);
            switch (FileName)
            {
                case "RemoveMyMenus.xml":
                case "MyMenus.xml":
                    return "";
                default:
                    SAPbouiCOM.Form oForm = SBO_Application.Forms.ActiveForm;
                    if (oForm.TypeEx == FileName.Replace(".xml", ""))
                        return oForm.UniqueID;
                    else
                        throw new Exception("Unable to open form");

                    //oForm.Resize(Convert.ToInt16(oForm.Width * iFontWidthRatio), Convert.ToInt16(oForm.Height * iFontHeightRatio));
                    
            }
        }

        public SAPbouiCOM.Form LoadFromXMLString(string sXml, int Xloc, int Yloc)
        {
            //Modify the form size to match the client Font Size
            sXml = ModifySize(sXml);

            //Modify Location of form.
            if (Xloc != 0)
                sXml = ModifyXLoc(sXml, Xloc);

            if (Yloc != 0)
                ModifyYLoc(sXml, Yloc);

            // load the form to the SBO application in one batch
            SBO_Application.LoadBatchActions(ref sXml);

            SAPbouiCOM.Form oForm = SBO_Application.Forms.ActiveForm;

            return oForm;

        }

        public SAPbouiCOM.Form LoadFromXMLFile(string FileName, int Xloc, int Yloc)
        {
            System.Xml.XmlDocument oXmlDoc = new System.Xml.XmlDocument();
            string sXml = "";
            // load the content of the XML File
            string sPath = System.Windows.Forms.Application.StartupPath.ToString();
            try
            {
                oXmlDoc.Load(sPath + "\\" + FileName);
            }
            catch (Exception ex)
            {
                if (eCommon.oEventLog != null) eCommon.oEventLog.WriteLine(DateTime.Now.ToString("yyyyMMdd HH:mm:ss") + Filler + "EX.MESSAGE : " + ex.Message);
            }
            sXml = oXmlDoc.InnerXml;


            return LoadFromXMLString(sXml, 0, 0);

        }



        public static string ModifyXLoc(string Original, int Xloc)
        {
            int leftCharStart = Original.IndexOf("left=\"") + 6;
            int leftCharEnd = Original.IndexOf("\"", leftCharStart);
            string Result = Original.Substring(0, leftCharStart) + Xloc + Original.Substring(leftCharEnd);
            return Result;
        }

        public static string ModifyYLoc(string Original, int Yloc)
        {
            int leftCharStart = Original.IndexOf("top=\"") + 5;
            int leftCharEnd = Original.IndexOf("\"", leftCharStart);
            string Result = Original.Substring(0, leftCharStart) + Yloc + Original.Substring(leftCharEnd);
            return Result;
        }

        /// <summary>
        /// This will take the XML string and modify the form left/width, top/height to match the current client font size
        /// </summary>
        /// <param name="Original">Original XML String</param>
        /// <returns></returns>
        public static string ModifySize(string Original)
        {

            System.Collections.Specialized.OrderedDictionary oTypes = new System.Collections.Specialized.OrderedDictionary();
            double dFontHeightRatio = Math.Round(eCommon.SBO_Application.GetFormItemDefaultHeight(BoFormSizeableItemTypes.fsit_EDIT) / 14.00, 2);        //Ratio is based on Edit text item. 14.00 is the reference Height that i created the forms in
            double dFontWidthRatio = Math.Round(eCommon.SBO_Application.GetFormItemDefaultWidth(BoFormSizeableItemTypes.fsit_EDIT) / 80.00, 2);                            //Ratio is based on Edit text item. 80.00 is the reference Width that i created the forms in

            oTypes.Add("left", dFontWidthRatio);
            oTypes.Add("width", dFontWidthRatio);
            oTypes.Add("top", dFontHeightRatio);
            oTypes.Add("height", dFontHeightRatio);

            foreach (string Type in oTypes.Keys)
            {
                int i = 0;
                double Ratio = (double)oTypes[Type];
                while (i < Original.Length)
                {
                    i = Original.IndexOf(Type + "=\"", i);
                    if (i > 0)
                    {
                        int iNextApos = Original.IndexOf("\"", i + Type.Length + 2);
                        string sContent = Original.Substring(i + Type.Length + 2, iNextApos - (i + Type.Length + 2));
                        int iContent = 0;
                        if (int.TryParse(sContent, out iContent))
                        {
                            Original = Original.Substring(0, i) + Type + "=\"" + Convert.ToInt16(iContent * Ratio).ToString() + Original.Substring(iNextApos);
                        }
                        i = iNextApos;
                    }
                    else
                    {
                        i = Original.Length;
                    }
                }
            }
            return Original;
        }



        /// <summary>
        /// Add a Query.
        /// </summary>
        /// <param name="sSql">The SQL statement to add</param>
        /// <param name="sQName">Query Name, will be created if not exists</param>
        /// <param name="sCategory">Category Name, will be created if not exists</param>
        /// <remarks> Add a saved query into SBO and Assign the Query to a formatted search. Query and Category will be created if not exists.</remarks>
        public static void AddQuery(string sSql, string sQName, string sCategory) //UserQueries 
        {
            SAPbobsCOM.QueryCategories oQC;
            SAPbobsCOM.Recordset oRS;
            oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery("Select CategoryId from OQCN WHERE CatName = '" + sCategory + "'");
            oQC = (SAPbobsCOM.QueryCategories)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQueryCategories);
            if (oRS.RecordCount > 0) oQC.Browser.Recordset = oRS;
            int iCatID = -1;

            if (oRS.RecordCount == 1)
            {
                iCatID = oQC.Code;
            }
            else if (oRS.RecordCount > 1)
            {
                throw new Exception("Multiple definition of category: [" + sCategory + "]!");
            }
            else
            {
                oQC = null;
                oQC = (SAPbobsCOM.QueryCategories)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQueryCategories);
                oQC.Name = sCategory;
                oQC.Permissions = "YYYYYYYYYYYYYYYYYYYY";
                iCatID = oQC.Add();
                if (iCatID != 0)
                {
                    throw new Exception("Query Category: " + oCompany.GetLastErrorDescription());
                }
                else
                {
                    iCatID = int.Parse(oCompany.GetNewObjectKey());
                }
            }


            SAPbobsCOM.UserQueries oUQ;
            oUQ = (SAPbobsCOM.UserQueries)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries);
            oRS.DoQuery("Select IntrnalKey FROM OUQR WHERE QCategory = '" + iCatID + "' AND QName = '" + sQName + "'");
            if (oRS.RecordCount > 0)
            {
                oUQ.Browser.Recordset = oRS;
            }
            string iQryID = "";
            if (oRS.RecordCount == 1)
            {
                iQryID = oUQ.InternalKey.ToString();
                string sNM = oQC.Name;
            }
            else if (oRS.RecordCount > 1)
            {
                throw new Exception("Multiple definition of Query: [" + sQName + "]!");
            }
            else
            {
                oUQ = null;
                oUQ = (SAPbobsCOM.UserQueries)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries);
                oUQ.Query = sSql;
                oUQ.QueryCategory = iCatID;
                oUQ.QueryDescription = sQName;
                iQryID = oUQ.Add().ToString();
                if (iQryID != "0")
                {
                    string sMsg = "";
                    int iErr = 0;
                    oCompany.GetLastError(out iErr, out  sMsg);
                    if (iErr != 0) throw new Exception("User Query: " + sMsg);
                }
                else
                {
                    string[] sKey;
                    sKey = (oCompany.GetNewObjectKey().Split('\t'));    //Split with Tab
                    iQryID = sKey[0];
                }
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUQ);
            oUQ = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oQC);
            oQC = null;
        }

        /// <summary>
        /// Add a formatted search.
        /// </summary>
        /// <param name="sSql">The SQL statement to add</param>
        /// <param name="sQName">Query Name, will be created if not exists</param>
        /// <param name="sCategory">Category Name, will be created if not exists</param>
        /// <param name="sFormID">The FormID to attach this Formatted Search</param>
        /// <param name="sItemID">Item which the FS is to be attached to</param>
        /// <param name="sFieldID">Item which the sItemID will be refreshed</param>
        /// <param name="sColID">The Column ID which the FS is to be attached. Incase sItemID is a matrix</param>
        /// <param name="bForceRefresh">Display Saved User-Defined Values</param>
        /// <remarks> Add a saved query into SBO and Assign the Query to a formatted search. Query and Category will be created if not exists.</remarks>
        public static void AddFS(string sSql, string sQName, string sCategory, string sFormID, string sItemID, string sFieldID, string sColID, bool bForceRefresh) //UserQueries 
        {
            SAPbobsCOM.QueryCategories oQC;
            SAPbobsCOM.Recordset oRS;
            if (sColID == null || sColID == "") sColID = "-1";

            oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery("Select CategoryId from OQCN WHERE CatName = '" + sCategory + "'");
            oQC = (SAPbobsCOM.QueryCategories)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQueryCategories);
            if (oRS.RecordCount > 0) oQC.Browser.Recordset = oRS;
            int iCatID = -1;
            if (oRS.RecordCount == 1)
            {
                iCatID = oQC.Code;
            }
            else if (oRS.RecordCount > 1)
            {
                throw new Exception("Multiple definition of category: [" + sCategory + "]!");
            }
            else
            {
                oQC = null;
                oQC = (SAPbobsCOM.QueryCategories)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQueryCategories);
                oQC.Name = sCategory;
                oQC.Permissions = "YYYYYYYYYYYYYYYYYYYY";
                iCatID = oQC.Add();
                if (iCatID != 0)
                {
                    throw new Exception("Query Category: " + oCompany.GetLastErrorDescription());
                }
                else
                {
                    iCatID = int.Parse(oCompany.GetNewObjectKey());
                }
            }


            SAPbobsCOM.UserQueries oUQ = null;

            oRS.DoQuery("Select IntrnalKey FROM OUQR WHERE QCategory = '" + iCatID + "' AND QName = '" + sQName + "'");
            if (oRS.RecordCount > 0)
            {
                oUQ = (SAPbobsCOM.UserQueries)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries);
                oUQ.Browser.Recordset = oRS;
            }
            String iQryID = "";
            if (oRS.RecordCount == 1)
            {
                iQryID = oUQ.InternalKey.ToString();
                string sNM = oQC.Name;
            }
            else if (oRS.RecordCount > 1)
            {
                throw new Exception("Multiple definition of Query: [" + sQName + "]!");
            }
            else
            {
                oUQ = null;
                oUQ = (SAPbobsCOM.UserQueries)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries);
                oUQ.Query = sSql;
                oUQ.QueryCategory = iCatID;
                oUQ.QueryDescription = sQName;
                iQryID = oUQ.Add().ToString();
                if (iQryID != "0")
                {
                    string sMsg = "";
                    int iErr = 0;
                    oCompany.GetLastError(out iErr, out  sMsg);
                    if (iErr != 0) { throw new Exception("User Query: " + sMsg); }
                }
                else
                {
                    string[] sKey;
                    sKey = (oCompany.GetNewObjectKey().Split('\t'));
                    iQryID = sKey[0];
                }
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUQ);
            oUQ = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oQC);
            oQC = null;

            if (sColID != "-1")
            {
                oRS.DoQuery("SELECT INDEXID FROM CSHS WHERE FORMID = '" + sFormID + "' AND ItemID = '" + sItemID + "' AND ColID = '" + sColID + "'");
            }
            else
            {
                oRS.DoQuery("SELECT INDEXID FROM CSHS WHERE FORMID = '" + sFormID + "' AND ItemID = '" + sItemID + "'");
            }

            if (!oRS.EoF)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                oRS = null;
                return;
            }

            //Link the Formatted search to the form, field
            SAPbobsCOM.FormattedSearches fs;
            fs = (SAPbobsCOM.FormattedSearches)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches);

            fs.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery;
            fs.QueryID = int.Parse(iQryID);
            fs.FormID = sFormID;
            fs.ItemID = sItemID;
            if (sColID != "-1") fs.ColumnID = sColID;
            if (bForceRefresh)
            {
                fs.FieldID = sFieldID;
                fs.Refresh = BoYesNoEnum.tYES;
                fs.ForceRefresh = BoYesNoEnum.tYES;
            }
            fs.ByField = BoYesNoEnum.tYES;

            int iAddErr = fs.Add();
            if (iAddErr != 0)
            {
                throw new Exception("Formatted Search: " + oCompany.GetLastErrorDescription());
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(fs);
            fs = null;
            return;
        }

        /// <summary>
        /// Set the default value for UDF
        /// </summary>
        /// <param name="Table"></param>
        /// <param name="FieldName"></param>
        /// <remarks></remarks>
        public static void SetDefaultUDFValue(string Table, string FieldName, string Value)
        {
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            oRS.DoQuery(String.Format("SELECT FieldID FROM CUFD WHERE TableID = '{0}' AND AliasID = '{1}' AND isnull(Dflt,'') = '' ", Table, FieldName));
            if (!oRS.EoF)
            {
                int FieldID = int.Parse(oRS.Fields.Item("FieldID").Value.ToString());
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                GC.Collect();
                SAPbobsCOM.UserFieldsMD oUFD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(BoObjectTypes.oUserFields);
                if (oUFD.GetByKey(Table, FieldID))
                {
                    if (oUFD.DefaultValue == "")
                    {
                        oUFD.DefaultValue = Value;
                        oUFD.Update();
                    }
                }
            }
            else
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                GC.Collect();
            }

        }

        public static void SetLinkedTable(string TableUID, string FieldUID, string LinkedTableUID, string LinkedUDOUID)
        {
            SAPbobsCOM.Recordset oRS = null;
            string sSQL = "";
            int FieldID = 0;
            try
            {
                sSQL = String.Format("SELECT RTable, FieldID FROM CUFD WHERE TableID = '{0}' AND AliasID = '{1}'", TableUID, FieldUID);
                oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRS.DoQuery(sSQL);
                if (!oRS.EoF)
                {
                    FieldID = int.Parse(oRS.Fields.Item("FieldID").Value.ToString());
                    if (oRS.Fields.Item("RTable").Value.ToString().Trim() == "")
                    {
                        ReleaseComObject((object)oRS);
                        GC.Collect();

                        SAPbobsCOM.UserFieldsMD oUFD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(BoObjectTypes.oUserFields);
                        if (oUFD.GetByKey(TableUID, FieldID))
                        {
                            
                            if (LinkedTableUID == "")
                                oUFD.LinkedUDO = LinkedUDOUID;
                            else
                                oUFD.LinkedTable = LinkedTableUID;

                            if (oUFD.Update() != 0)
                            {
                                throw new Exception(oCompany.GetLastErrorDescription());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(String.Format("Unable to set linked table to UDF '{0}' in Table '{1}'. Please update manually. {2}", FieldUID, TableUID, ex.Message), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                ReleaseComObject((object)oRS);
            }
        }


        #endregion




        public static void BrowseForm(string mnuUID, string ItemUID, string SearchCode)
        {
            string FormUID = DrawForm(mnuUID + ".xml", null);
            SAPbouiCOM.Form Form = SBO_Application.Forms.Item(FormUID);
            bool bSuccess = false;
            Form.Freeze(true);
            try
            {
                Form.Mode = BoFormMode.fm_FIND_MODE;
                Form.Items.Item(ItemUID).Enabled = true;
                ((SAPbouiCOM.EditText)Form.Items.Item(ItemUID).Specific).String = SearchCode;
                Form.Items.Item("1").Click(BoCellClickType.ct_Regular);
                bSuccess = true;
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Unable to link. " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                Form.Freeze(false);
                if (!bSuccess) Form.Close();
            }
        }

        public static void ReleaseComObject(object oObject)
        {
            if (oObject != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oObject);
                oObject = null;
            }
        }

        public static bool AddUDF(string TableName, string FieldName, string Description, SAPbobsCOM.BoFieldTypes FieldType, int EditSize, string LinkedTable)
        {
            if (LinkedTable == null) LinkedTable = "";
            SAPbobsCOM.UserFieldsMD oUserFieldsMD;
            int lRetCode = 0;
            string sErrMsg = "";

            oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

            // Setting the Field's mandatory properties
            oUserFieldsMD.TableName = TableName;
            oUserFieldsMD.Name = FieldName;
            oUserFieldsMD.Description = Description;
            oUserFieldsMD.Type = FieldType;
            if (FieldType != SAPbobsCOM.BoFieldTypes.db_Date)
            {
                oUserFieldsMD.EditSize = EditSize;
            }

            if (LinkedTable.Length > 0)
            {
                if (LinkedTable.Contains("#Predefined"))
                {
                    string[] Fields = LinkedTable.Split(',');
                    string[] ValidValues;
                    int iInd = 0;
                    foreach (string sField in Fields)
                    {
                        if (iInd == 0)
                        {
                            iInd += 1;
                            continue;
                        }
                        ValidValues = sField.Split(':');
                        if (iInd > 2)
                        {
                            oUserFieldsMD.ValidValues.Add();
                        }
                        oUserFieldsMD.ValidValues.Value = ValidValues[0].ToString();
                        oUserFieldsMD.ValidValues.Description = ValidValues[1].ToString();
                        iInd += 1;
                    }
                }
                else
                {
                    oUserFieldsMD.LinkedTable = LinkedTable;
                }
            }


            // Adding the Field to the Table
            lRetCode = oUserFieldsMD.Add();

            // Check for errors
            if (lRetCode != 0)
            {
                oCompany.GetLastError(out lRetCode, out  sErrMsg);
                SBO_Application.MessageBox("Unable to add UDF " + TableName + "\\" + FieldName + "\n\r" + sErrMsg, 1, "OK", null, null);
                oEventLog.WriteLine(DateTime.Now.ToString("yyyyMMdd HH:mm:ss") + Filler + "Unable to add UDF : " + TableName + "\\" + FieldName + sErrMsg);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);

                return false;
            }
            else
            {
                oEventLog.WriteLine("Field: '" + oUserFieldsMD.Name + "' was added successfuly to " + oUserFieldsMD.TableName + " Table");
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                return true;
            }
        }

        public static void OpenSystemForm(string ObjectType, string Code)
        {
            SAPbouiCOM.Form oForm = SBO_Application.Forms.Add("", BoFormTypes.ft_Fixed, -1);
            SAPbouiCOM.Item oItem;
            oForm.DataSources.UserDataSources.Add("txtCode", BoDataType.dt_SHORT_TEXT, 50);
            oItem = oForm.Items.Add("txtCode", BoFormItemTypes.it_EDIT);
            SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oItem.Specific;
            oEdit.DataBind.SetBound(true, "", "txtCode");
            oItem = oForm.Items.Add("lnkCode", BoFormItemTypes.it_LINKED_BUTTON);
            SAPbouiCOM.LinkedButton oLnk = (SAPbouiCOM.LinkedButton)oItem.Specific;
            oEdit.String = Code;
            oLnk.LinkedObjectType = ObjectType;
            oItem.LinkTo = "txtCode";

            oItem.Click(BoCellClickType.ct_Regular);

            oForm.Close();

        }


        public static string GetTableName(string ObjectCode)
        {
            switch (ObjectCode)
            {
                case "13": return "INV";
                case "14": return "RIN";
                case "15": return "DLN";
                case "18": return "PCH";
                case "19": return "RPC";
                case "20": return "PDN";
                case "21": return "RPD";
                case "22": return "POR";
                case "23": return "QUT";
                case "17": return "RDR";
                case "46": return "VPM";
                case "59": return "IGN";
                case "60": return "IGE";
                case "67": return "WTR";
                case "203": return "DPI";
                case "204": return "DPO";
                case "540000006": return "PQT";
                default: return ObjectCode;
            }


        }

        public static string GetObjectName(string ObjectCode)
        {
            switch (ObjectCode)
            {
                case "13": return "AR Invoice";
                case "14": return "AR CN";
                case "15": return "Delivery Note";
                case "18": return "AP Invoice";
                case "19": return "AP CN";
                case "20": return "PO Receipt";
                case "21": return "PO G : return";
                case "22": return "Purchase Order";
                case "23": return "Quotation";
                case "17": return "Sales Order";
                case "46": return "Outgoing Payment";
                case "59": return "Inv G Receipt";
                case "60": return "Inv G Issue";
                case "67": return "OWTR";
                case "203": return "AR Down Payment";
                case "204": return "AP Down Payment";
                default: return ObjectCode;
            }
        }

        public static string GetDocNum(int DocEntry, string DocObjectCode)
        {
            string sDocNum = "";
            SAPbobsCOM.Documents oDoc = null;
            try
            {
                oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)int.Parse(DocObjectCode));
                if (oDoc.GetByKey(DocEntry))
                {
                    sDocNum = oDoc.DocNum.ToString();
                }
            }
            catch { }
            finally
            {
                ReleaseComObject((object)oDoc);
            }
            return sDocNum;
        }

        public static bool isManagedByBatch(string ItemCode)
        {
            bool isBatch = false;
            SAPbobsCOM.Items oItm = null;
            try
            {
                oItm = (SAPbobsCOM.Items)oCompany.GetBusinessObject(BoObjectTypes.oItems);
                if (oItm.GetByKey(ItemCode))
                {
                    if (oItm.ManageBatchNumbers == BoYesNoEnum.tYES)
                    {
                        isBatch = true;
                    }
                }
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.ToString(), 1, "OK", null, null);
                isBatch = false;
            }
            finally
            {
                ReleaseComObject((Object)oItm);
            }
            return isBatch;
        }

        public static bool UpdateDimensions(int Code, string Name)
        {
            SAPbobsCOM.CompanyService oCmpSrv;
            SAPbobsCOM.DimensionsService oDIMService;
            SAPbobsCOM.DimensionParams oDIMParams;
            SAPbobsCOM.Dimension oDIM;

            oCmpSrv = oCompany.GetCompanyService();
            oDIMService = (SAPbobsCOM.DimensionsService)oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.DimensionsService);
            oDIMParams = (SAPbobsCOM.DimensionParams)oDIMService.GetDataInterface(SAPbobsCOM.DimensionsServiceDataInterfaces.dsDimensionParams);

            oDIMParams.DimensionCode = Code;

            try
            {
                oDIM = oDIMService.GetDimension(oDIMParams);
            }
            catch
            {
                return false;
            }


            if (oDIM.DimensionDescription != Name)
            {
                oDIM.IsActive = BoYesNoEnum.tYES;
                oDIM.DimensionDescription = Name;
                try
                {
                    oDIMService.UpdateDimension(oDIM);
                    return true;
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox("Unable to create Dimensions for " + Name + "\r\n" + ex.Message, 1, "OK", null, null);
                    return false;
                }
            }
            else
            {
                return true;
            }

        }

        public static void UpdateRoles(string RoleCode, string RoleDesc)
        {
            SAPbobsCOM.CompanyService oCompSrv = oCompany.GetCompanyService();
            SAPbobsCOM.EmployeeRolesSetupService oRoleSrv = (SAPbobsCOM.EmployeeRolesSetupService)oCompSrv.GetBusinessService(ServiceTypes.EmployeeRolesSetupService);
            EmployeeRoleSetup RoleSetup = null;

            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                oRS.DoQuery("SELECT * FROM OHTY WHERE Name = '" + RoleCode + "'");

                if (oRS.EoF)
                {
                    RoleSetup = (EmployeeRoleSetup)oRoleSrv.GetDataInterface(EmployeeRolesSetupServiceDataInterfaces.erssEmployeeRoleSetup);

                    RoleSetup.Name = RoleCode;
                    RoleSetup.Description = RoleDesc;
                    oRoleSrv.AddEmployeeRoleSetup(RoleSetup);
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                ReleaseComObject((object)oRS);
                ReleaseComObject((object)oCompSrv);
                ReleaseComObject((object)oRoleSrv);
                ReleaseComObject((object)RoleSetup);
            }
        }

        public static void UpdateItemProperty(int PropNumber, string PropName)
        {
            SAPbobsCOM.Recordset oRS = null;
            SAPbobsCOM.ItemProperties ItmProp = null;
            string sSQL = "";

            try
            {
                sSQL = String.Format("SELECT * FROM OITG WHERE ItmsTypCod = {0} AND ItmsGrpNam = '{1}'", PropNumber, "Items Property " + PropNumber);
                oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRS.DoQuery(sSQL);

                if (!oRS.EoF)
                {
                    ItmProp = (SAPbobsCOM.ItemProperties)oCompany.GetBusinessObject(BoObjectTypes.oItemProperties);
                    if (ItmProp.GetByKey(PropNumber))
                    {
                        ItmProp.PropertyName = PropName;
                        ItmProp.Update();
                    }
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                ReleaseComObject((object)oRS);
                ReleaseComObject((object)ItmProp);
            }
        }

        private static string GetSegmentedAccountCode(string SegmentedAccountCode)
        {
            string sStr = "";
            SAPbobsCOM.Recordset oRS;
            SAPbobsCOM.SBObob oBOB;
            SAPbobsCOM.ChartOfAccounts oAcct;
            oAcct = (SAPbobsCOM.ChartOfAccounts)oCompany.GetBusinessObject(BoObjectTypes.oChartOfAccounts);
            oBOB = (SAPbobsCOM.SBObob)oCompany.GetBusinessObject(BoObjectTypes.BoBridge);
            oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            // When working with segmentation use this function
            // to find the account key in the ChartOfAccount object

            oRS = oBOB.GetObjectKeyBySingleValue(BoObjectTypes.oChartOfAccounts, "FormatCode", SegmentedAccountCode.Replace(sAcctSeparator, ""), BoQueryConditions.bqc_Equal);

            //The Recordset retrieves the value of the key (for example, sStr = _SYS00000000010).
            if (oRS.RecordCount > 0)
            {
                sStr = (string)oRS.Fields.Item(0).Value;
            }
            else
            {
                throw new Exception("Unable to get the account code with the segmented code '" + SegmentedAccountCode + "'");
            }

            return sStr;
        }


        private static bool CheckLicenceEX(string AddOnName)
        {
            try
            {
                TWM_Licence.TWM_Licence oLic = new TWM_Licence.TWM_Licence(eCommon.SBO_Application,eCommon.oCompany, AddOnName, new Byte[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 10, 73, 1, 5, 75, 1, 8 });
                if (!oLic.IsValid)
                {
                    eCommon.SBO_Application.StatusBar.SetText("Could not start addon. " + oLic.LastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    eCommon.oCompany.Disconnect();
                    //Environment.Exit(1);
                    return false;
                }
                else
                {
                    if (oLic.DaysToExpiry < 10)
                    {
                        if (oLic.DaysToExpiry > 0)
                        {
                            eCommon.SBO_Application.MessageBox("Your add on " + AddOnName + " will expire in " + oLic.DaysToExpiry + " days. Please contact support for the license.", 1, "OK", null, null);
                        }
                        else
                        {
                            eCommon.SBO_Application.MessageBox("Your add on " + AddOnName + " expires today. Please contact support for the license.", 1, "OK", null, null);
                        }
                    }
                    return true;
                }
            }
            catch
            {
                eCommon.SBO_Application.StatusBar.SetText("Could not start addon. No licence found.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                eCommon.oCompany.Disconnect();
                System.Windows.Forms.Application.Exit();
                return false;
            }
        }


        private static bool IsDesignTime()
        {
            try
            {
                System.Diagnostics.Process cProcess = System.Diagnostics.Process.GetCurrentProcess();
                if (cProcess.ProcessName.Contains(".vshost")) //This part is for DesignTime Mode
                    return true;
                else
                    return false;
            }
            catch
            {
                return false;
            }
        }


        public static SAPbobsCOM.Activity GetActivityByKey(int Code)
        {
            SAPbobsCOM.CompanyService oCmpSrv = null;
            SAPbobsCOM.ActivitiesService oActService = null;
            SAPbobsCOM.ActivityParams oActParam = null;
            SAPbobsCOM.Activity oAct = null;




            oCmpSrv = oCompany.GetCompanyService();
            oActService = (SAPbobsCOM.ActivitiesService)oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ActivitiesService);

            oActParam = (SAPbobsCOM.ActivityParams)oActService.GetDataInterface(ActivitiesServiceDataInterfaces.asActivityParams);
            oActParam.ActivityCode = Code;

            try
            {
                oAct = oActService.GetActivity(oActParam);
            }
            catch
            {
                oAct = null;
            }

            return oAct;
        }



        public static System.Data.DataTable OpenXls(string sFileName, string sSheetName, string SelectQuery)
        {
            System.IO.FileInfo f = new System.IO.FileInfo(sFileName);
            string sXLconStr = "";
            System.Data.OleDb.OleDbConnection conXL = new System.Data.OleDb.OleDbConnection();
            System.Data.DataTable dt = new System.Data.DataTable(); ;
            System.Data.OleDb.OleDbDataAdapter da;
            string sSQL = "";

            switch (f.Extension.ToUpper())
            {
                case ".XLS":
                    if (oEventLog != null) oEventLog.WriteLine(DateTime.Now + Filler + "Opening .XLS file : " + sFileName);
                    sXLconStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sFileName + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1;\"";
                    break;
                case ".XLSX":
                    if (oEventLog != null) oEventLog.WriteLine(DateTime.Now + Filler + "Opening .XLSX file : " + sFileName);
                    sXLconStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sFileName + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1;\"";
                    break;
            }
            conXL.ConnectionString = sXLconStr;

            if (SelectQuery != null && SelectQuery.Trim() != "")
            {
                sSQL = SelectQuery + " FROM [" + sSheetName + "]";
            }
            else
            {
                sSQL = "SELECT  * FROM [" + sSheetName + "]";
            }

            try
            {
                conXL.Open();
                da = new System.Data.OleDb.OleDbDataAdapter(sSQL, conXL);
                if (oEventLog != null) oEventLog.WriteLine(DateTime.Now + Filler + "Read all file content");
                da.Fill(dt);
                da.Dispose();
            }
            catch (Exception ex)
            {
                if (oEventLog != null) oEventLog.WriteLine(DateTime.Now + Filler + "EX.MESSAGE : " + ex.Message);
                SBO_Application.MessageBox("Cannot open source file \r\n" + ex.Message, 1, "OK", null, null);
            }



            conXL.Close();
            conXL.Dispose();
            if (oEventLog != null) oEventLog.WriteLine(DateTime.Now + Filler + "File read successful");


            dt.TableName = sSheetName;
            string sLine = "";
            foreach (System.Data.DataColumn col in dt.Columns)
            {
                sLine += col.ColumnName + "     ";
            }
            System.Diagnostics.Debug.WriteLine(sLine);
            sLine = "";
            foreach (System.Data.DataRow row in dt.Rows)
            {
                foreach (System.Data.DataColumn columns in dt.Columns)
                {
                    sLine += row[columns.ColumnName] + "     ";
                }
                System.Diagnostics.Debug.WriteLine(sLine);
                sLine = "";
            }


            return dt;
        }
        public static System.Data.DataTable GetWorksheets(string sFilename)
        {
            System.IO.FileInfo f = new System.IO.FileInfo(sFilename);
            string sXLconStr = "";
            System.Data.OleDb.OleDbConnection conXL = new System.Data.OleDb.OleDbConnection();
            System.Data.DataTable dt = new System.Data.DataTable();
            System.Windows.Forms.Application.UseWaitCursor = true;
            if (oEventLog != null) oEventLog.WriteLine(DateTime.Now + Filler + "Extracting worksheets from " + sFilename);
            switch (f.Extension.ToUpper())
            {
                case ".XLS":
                    sXLconStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sFilename + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
                    break;
                case ".XLSX":
                    sXLconStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sFilename + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1;\"";
                    break;
            }
            conXL.ConnectionString = sXLconStr;
            try
            {
                conXL.Open();
                dt = conXL.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new Object[] { null, null, null, "TABLE" });
                if (oEventLog != null) oEventLog.WriteLine(DateTime.Now + Filler + "Worksheets retrieved : " + dt.Rows.Count);
                conXL.Close();
            }
            catch (Exception ex)
            {

                if (oEventLog != null) oEventLog.WriteLine(DateTime.Now + Filler + "EX.MESSAGE : " + ex.Message);
                SBO_Application.MessageBox("Could not open source file \r\n" + ex.Message, 1, "OK", null, null);
            }

            System.Windows.Forms.Application.UseWaitCursor = false;
            return dt;
        }

        /// <summary>
        /// Add a userkey index for the table. The Fields Name must exclude the 'U_'
        /// </summary>
        /// <param name="sTableName">include the '@' prefix</param>
        /// <param name="sKeyName"></param>
        /// <param name="Fields">Exclude the 'U_' prefix</param>
        /// <param name="Unique"></param>
        /// <returns></returns>
        public static bool AddUserKey(string sTableName, string sKeyName, string[] Fields, bool Unique)
        {
            //Check whether the key is already added.
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRS.DoQuery(string.Format("SELECT 1 FROM OUKD WHERE TableName = '{0}' AND KeyName='{1}'", sTableName, sKeyName));
                if (!oRS.EoF) return true;
            }
            finally
            {
                ReleaseComObject(oRS);
            }

            SAPbobsCOM.UserKeysMD oMD = null;
            try
            {
                oMD = (SAPbobsCOM.UserKeysMD)oCompany.GetBusinessObject(BoObjectTypes.oUserKeys);
                oMD.TableName = sTableName;
                oMD.KeyName = sKeyName;

                bool bFirst = true;
                foreach (string sKey in Fields)
                {
                    if (!bFirst)
                    {
                        oMD.Elements.Add();
                    }
                    else
                    {
                        bFirst = false;
                    }

                    oMD.Elements.ColumnAlias = sKey;
                }

                if (Unique) oMD.Unique = BoYesNoEnum.tYES;

                if (oMD.Add() != 0)
                {
                    throw new Exception(oCompany.GetLastErrorDescription());
                }

                return true;
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(string.Format("Unable to add key index. [{0}.{1}]-{2}", sTableName, Fields[0], ex.Message), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
                return false;
            }
            finally
            {
                ReleaseComObject(oMD);
            }
        }

        public static void AddSP(string SPName, string sSQL)
        {
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRS.DoQuery(string.Format("SELECT ISNULL(OBJECT_ID('{0}'),-1)", SPName));
                if (oRS.Fields.Item(0).Value.ToString() == "-1")
                {
                    oRS.DoQuery(sSQL);
                }
            }
            catch
            {
                throw new Exception(string.Format("'{0}' not found. Please contact support", SPName));
            }
            finally
            {
                ReleaseComObject(oRS);
            }
        }

        public static string GetAdvAcctDetermintaion(string ItemCode, string WhsCode, DateTime RefDate, string BPGrpCod, string ShipCountr, string ShipState
    , string LicTradNum, string VatGroup, string CardCode, string CmpPrivate)
        {
            string sSQL = string.Format(@"
                DECLARE @ItemCode NVARCHAR(20)='{0}'
                DECLARE @WhsCOde NVARCHAR(8)='{1}'
                DECLARE @RefDate DATE= '{2:yyyyMMdd}'
                DECLARE @BPGrpCod INT = {3}
                DECLARE @ShipCountr NVARCHAR(3) = '{4}'
                DECLARE @ShipState NVARCHAR(3) = '{5}'
                DECLARE @LicTradNum NVARCHAR(32) ='{6}'
                DECLARE @VatGroup NVARCHAR(8) = '{7}'
                DECLARE @CardCode NVARCHAR(15) = '{8}'
                DECLARE @CmpPrivate NVARCHAR(1) = '{9}'

                DECLARE @NewAcctDe NVARCHAR(1)
                SELECT @NewAcctDe=NewAcctDe FROM OADM

                IF @NewAcctDe='N' 
                BEGIN
	                SELECT (CASE glmethod   WHEN 'L' THEN t1.WipAcct  WHEN 'W' THEN t2.WipAcct  WHEN 'C' THEN t3.WipAcct END) WIPAcctCode  
	                FROM oitm t0    
	                INNER JOIN oitw t1 ON t0.itemcode=t1.itemcode    
	                LEFT JOIN owhs t2 ON t1.whscode=t2.whscode    
	                LEFT JOIN oitb t3 ON t0.itmsgrpcod=t3.itmsgrpcod   
	                WHERE T0.ItemCode = @ItemCode AND T1.WhsCode = @WhsCode
                END
                ELSE
                BEGIN
                /*Advance GL Account Determination*/

                DECLARE @SQL NVARCHAR(MAX)='
                DECLARE @AbsEntry INT =0
                DECLARE @Acct NVARCHAR(15)=''''
                SELECT @AbsEntry  =MAX(AbsEntry) FROM OACP WHERE F_RefDate <=@RefDate
                SELECT TOP 1 @Acct = T1.WipAcct
                FROM OITM T0 JOIN OGAR T1 ON (T0.ItemCode = T1.ItemCode OR T1.ItemCode=''!^|'') 
	                AND ISNULL(T1.WipAcct,'''')<>''''
	                AND (T0.ItmsGrpCod = T1.ItmsGrpCod OR T1.ItmsGrpCod = -1) 
	                AND (T1.WhsCode = @WhsCode OR T1.WhsCode = ''!^|'')
	                AND T1.Active = ''Y''
	                AND @RefDate BETWEEN T1.F_RefDate AND T1.T_RefDate
	                AND @RefDate BETWEEN ISNULL(T1.FromDate,''19010101'') AND ISNULL(T1.ToDate,''21991231'')
	                AND (T0.GLPickMeth = T1.GLMethod  OR T1.GLMethod=''A'')
	                AND (T1.BPGrpCod = @BPGrpCod OR T1.BPGrpCod=-1) 
	                AND (T1.ShipCountr = @ShipCountr OR T1.ShipCountr= ''!^|'')
	                AND (T1.ShipState = @ShipState OR T1.ShipState= ''!^|'')
	                AND (T1.LicTradNum = @LicTradNum OR T1.LicTradNum = ''!^|'')
	                AND (T1.VatGroup = @VatGroup OR T1.VatGroup = ''!^|'')
	                AND (T1.CardCode = @CardCode OR T1.CardCode = ''!^|'')
	                AND (T1.CmpPrivate = @CmpPrivate OR T1.CmpPrivate = ''!^|'')
                WHERE T0.ItemCode = @ItemCode'

                DECLARE @Order NVARCHAR(2000)
                SELECT @Order=(SELECT 
	                CASE DmcAlias
		                WHEN 'Item Group' THEN 'T1.ItmsGrpCod'
		                WHEN 'Item Code' THEN 'T1.ItemCode'
		                WHEN 'Warehouse Code' THEN 'T1.WhsCode'
		                WHEN 'Business Partner Group' Then 'T1.BPGrpCod'
		                WHEN 'Ship-to Country' Then 'T1.ShipCountr'
		                WHEN 'Ship-to State' Then 'T1.ShipState'
		                WHEN 'Federal Tax ID' Then 'T1.LicTradNum'
		                WHEN 'Tax Code' Then 'T1.VatGroup'
		                WHEN 'BP Code' Then 'T1.CardCode'
		                WHEN 'BP Type' Then 'T1.CmpPrivate'
	                END + ', '
                FROM ODMC
                WHERE Active = 'Y'
                ORDER BY Priority
                FOR XML PATH('')) + '1'

                SET @SQL =@SQL + ' ORDER BY ' + @Order

                SET @SQL = @SQL + ' ' + '

                if @Acct=''''
	                SELECT @Acct = WipAcct From OACP WHERE AbsEntry = @AbsEntry


                SELECT @Acct
                '	
                --SELECT @SQL

                Exec sp_ExecuteSQL @SQL,N'@RefDate DATE, @ItemCode NVARCHAR(20), @WhsCode NVARCHAR(15), @BPGrpCod INT 
                , @ShipCountr NVARCHAR(3), @ShipState NVARCHAR(3), @LicTradNum NVARCHAR(32), @VatGroup NVARCHAR(8), @CardCode NVARCHAR(15)
                , @CmpPrivate NVARCHAR(1)', @RefDate, @ItemCode, @WhsCOde, @BPGrpCod, @ShipCountr, @ShipState, @LicTradNum, @VatGroup, @CardCode, @CmpPrivate
	
                END", ItemCode, WhsCode, RefDate, BPGrpCod, ShipCountr, ShipState, LicTradNum, VatGroup, CardCode, CmpPrivate);
            return ExecuteScalar(sSQL).ToString();


        }


        /// <summary>
        /// Determine if a Portable Executable is of 32 or 64bit
        /// </summary>
        /// <param name="pFilePath">The file path</param>
        /// <returns>0x10b - PE32 or 0x20b - PE32+</returns>
        public static ushort GetPEArchitecture(string pFilePath)
        {
            ushort architecture = 0;
            try
            {
                using (System.IO.FileStream fStream = new System.IO.FileStream(pFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                {
                    using (System.IO.BinaryReader bReader = new System.IO.BinaryReader(fStream))
                    {
                        if (bReader.ReadUInt16() == 23117) //check the MZ signature
                        {
                            fStream.Seek(0x3A, System.IO.SeekOrigin.Current); // seek to e_lfanew.
                            fStream.Seek(bReader.ReadUInt32(), System.IO.SeekOrigin.Begin); //Seek to the start of the NT header.
                            if (bReader.ReadUInt32() == 17744) // check the PE\0\0 signature.
                            {
                                fStream.Seek(20, System.IO.SeekOrigin.Current); // seek past the file header, and
                                architecture = bReader.ReadUInt16(); // read the magic number of the optional header.
                            }
                        }
                    }
                }
            }
            catch (Exception) { /* todo: Any exception handling you want to do, personally I just take 0 as a sign of failure */}
            //if architecture returns 0, there has been an error.
            return architecture;
        }

        //Need to Imports System.Management ( Add Reference )
        private static int getProcessParentID(string cName, int cID)
        {
            SelectQuery query = new SelectQuery("SELECT * FROM Win32_Process WHERE Name like '" + cName + ".exe' and ProcessId = " + cID);
            ManagementObjectSearcher mgmtSearcher = new ManagementObjectSearcher(query);
            int kRet = -1;
            foreach (ManagementObject p in mgmtSearcher.Get())
            {
                string[] s = new string[1];
                p.InvokeMethod("GetOwner", (Object[])s);
                // Source Code link : http://www.vbdotnetforums.com/windows-services/4022-kill-specific-process.html
                // More Object Reference at this link : http://msdn.microsoft.com/en-us/library/aa394372(VS.85).asp
                kRet = int.Parse(p["ParentProcessId"].ToString());
            }
            return kRet;
        }


        /// <summary>
        /// This procedure checks for the same running instance on local machine.
        /// killFlag determine whether to kill the other instance.
        /// </summary>
        /// <param name="killFlag">Whether to kill the other instance.</param>
        /// <returns>Returns FALSE if there's an running process</returns>
        private static bool checkInstance(bool killFlag)  //Return Value : Returns FALSE if there's an running process
        {
            bool BufferFlag = true;
            System.Diagnostics.Process cProcess = System.Diagnostics.Process.GetCurrentProcess();
            System.Diagnostics.Process[] aProcesses = System.Diagnostics.Process.GetProcessesByName(cProcess.ProcessName);
            //aProcesses = Process.GetProcessesByName("TWM_KKLE_Trucking")
            int cParentID = getProcessParentID(cProcess.ProcessName, cProcess.Id);
            int xParentID = 0;

            foreach (System.Diagnostics.Process xProcess in aProcesses)
            {
                if (xProcess.Id != cProcess.Id)     //ignore the current (self)
                {
                    if (System.Reflection.Assembly.GetExecutingAssembly().Location == cProcess.MainModule.FileName) //'Check the running process with same EXE 
                    {
                        xParentID = getProcessParentID(xProcess.ProcessName, xProcess.Id);
                        if (xParentID == cParentID)
                        {
                            if ((bool)killFlag)
                            {
                                xProcess.Kill();
                                //MessageBox.Show("New / Parent = " & cProcess.Id & " / " & cParentID & " Old / Parent = " & xProcess.Id & " / " & xParentID, " Old Application was killed ")
                                MessageBox.Show("Running Addon for the same instance of SAP was terminated.", "old Process Killed", MessageBoxButtons.OK);
                                BufferFlag = true;
                            }
                            else
                            {
                                //If MessageBox.Show("New / Parent = " & cProcess.Id & " / " & cParentID & " Old / Parent = " & xProcess.Id & " / " & xParentID, " Wanna Kill Old Application ?", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                                if (MessageBox.Show("Found Same Addon was running for the same instance of SAP, wan to terminate the old ?", "wanna kill old process ? ", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                {
                                    xProcess.Kill();
                                    BufferFlag = true;
                                }
                                else
                                {
                                    MessageBox.Show("Application is already running", "Program Terminated!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                    BufferFlag = false;
                                }
                            }
                            //Exit For
                        }
                        else
                        {
                            //If xParentID points to non existent process, kill the process also.
                            try
                            {
                                System.Diagnostics.Process xPProcess = System.Diagnostics.Process.GetProcessById(xParentID);
                            }
                            catch
                            {
                                //Parent Process is not running. Kill it
                                xProcess.Kill();
                            }
                        }
                    }
                }
            }
            return BufferFlag;
        }

        public static string ConvertBase(string s, int FromBase, int ToBase)
        {
            string ConvertBases;
            //  Convert number in string representation in fromBase into toBase. Return result as a string
            //  Return error if input is empty
            if (String.IsNullOrEmpty(s)) return "";
            //  only do base 2 to base 36 (digit represented by charecaters 0-Z)"
            if ((FromBase < 2 || FromBase > 36) || (ToBase < 2 || ToBase > 36)) return "";
            s = s.ToUpper();  //  Convert to uppercase
            const string Allowed = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            if (System.Text.RegularExpressions.Regex.IsMatch(s, "^[" + Allowed.Substring(0, FromBase) + "]$*") == false) return "";
            // convert string to an array of integer digits representing number in frombase
            int il = s.Length;
            int[] fs; fs = new int[il];

            for (int i = s.Length - 1; i >= 0; i--) { fs[s.Length - (i + 1)] = Allowed.IndexOf(s[i]); }
            int ol = il * (FromBase / ToBase + 1);  // find how many digits the output needs
            int[] acc;
            acc = new int[ol + 10]; // assign accumulation array
            int[] Result;
            Result = new int[ol + 10]; // assign the result array
            acc[0] = 1; // initialise first acculamtion array element with number 1 
            int ip = 0;
            // for each input digit
            for (int i = 0; i <= il; i++)
            {
                for (int j = 0; j <= ol; j++) // add the input digit times (fromBase^i in baseTo) to the output accumulator
                {
                    Result[j] += acc[j] * fs[i];
                    ip = j;
                    while (Result[ip] >= ToBase)  // fix & cascade any which exceeds toBase
                    {
                        Result[ip + 1] += Result[ip] / ToBase;
                        Result[ip] = Result[ip] % ToBase;
                        ip++;
                    }
                }
                // Calculate the next power from^i) in toBase format
                for (int j = 0; j <= ol; j++) { acc[j] *= FromBase; }
                ip = 0;
                while (acc[ip] >= ToBase) //check for any which exceed toBase 
                {
                    acc[ip + 1] += acc[ip] / ToBase;
                    acc[ip] = acc[ip] % ToBase;
                    ip++;
                }
            }
            // convert the output to string format (digits 0,toBase-1 converted to 0-Z characters) 
            ConvertBases = String.Empty; // initialise output string
            ip = ol;
            while (Result[ip] == 0)
            {
                ip--;
            }
            while (ip >= 0)
            {
                ConvertBases += Allowed[Result[ip]];
                ip--;
            }
            if (String.IsNullOrEmpty(ConvertBases)) return "0";  //input was zero, return 0
            // return the converted string
            return ConvertBases;
        }


        /// <summary>
        /// Returns the number of minutes represented in SAPTime
        /// </summary>
        /// <param name="SAPTime">A 24hours format of time without any separator. ie 1300, 700(7 AM), 50(00:50 AM), 1(00:01 AM)</param>
        /// <returns>Int. Number of minutes from 00:00 AM</returns>
        public static int ConvertIntoMinutes(string SAPTime)
        {
            string sHourPart = "";
            string sMinutePart = "";
            int iResult = 0;
            if (!int.TryParse(SAPTime, out iResult))
            {
                throw new Exception("Time is not numeric");
            }

            if (SAPTime.Length == 3)
            {
                sHourPart = SAPTime.Substring(0, 1);
                sMinutePart = SAPTime.Substring(1, 2);
            }
            else if (SAPTime.Length == 4)
            {
                sHourPart = SAPTime.Substring(0, 2);
                sMinutePart = SAPTime.Substring(2, 2);
            }
            else if (SAPTime.Length < 3)
            {
                sHourPart = "0";
                sMinutePart = SAPTime;
            }

            return int.Parse(sHourPart) * 60 + int.Parse(sMinutePart);


        }

        /// <summary>
        /// Returns an 24hour integer representation of the time input
        /// </summary>
        /// <param name="Time">Time string should be in this format (hh:mmAM/hh:mmPM)</param>
        /// <returns></returns>
        public static int ConvertTimeToInt(string Time)
        {
            string sTemp = "";
            int iResult = 0;

            //If time is already in Numeric, return it directly
            if (int.TryParse(Time, out iResult)) return iResult;
            //If time does not end with AM/PM, but contains the ':', strip the ':' and return it directly
            switch (Time.Substring(Time.Length - 2).ToUpper())
            {
                case "AM":
                case "PM":
                    //do nothing.. handle it after select case
                    break;
                default:
                    //Time is in format of HH:mm, no AM/PM
                    sTemp = Time.Replace(":", "");
                    if (int.TryParse(sTemp, out iResult))
                        return iResult;
                    else
                        return 0;
            }

            Time = ("0" + Time);
            Time = Time.Substring(Time.Length - 7);
            string AMPM = Time.Substring(Time.Length - 2);
            string sTime = Time.Substring(0, 5);
            sTime = sTime.Replace(":", "");
            int iTime = 0;
            if (AMPM == "AM" && sTime.Substring(0, 2) == "12")
            {
                sTime = "00" + sTime.Substring(2, 2);
            }

            if (int.TryParse(sTime, out iResult))
                iTime = iResult;
            else
                return 0;


            if (AMPM == "PM" && iTime < 1200)
                iTime += 1200;


            return iTime;
        }

        /// <summary>
        /// Convert an 24hour integer representation of time to Time String
        /// </summary>
        /// <param name="Time">Integer. A 24hours format of time without any separator. ie 1300, 700(7 AM), 50(00:50 AM), 1(00:01 AM) </param>
        /// <returns>Formatted string of time in hh:mmAM/PM</returns>
        public static string ConvertTimeToString(int Time)
        {
            //Time string should be in this format (hh:mmAM/hh:mmPM)
            string sTime = "";
            Time = Math.Abs(Time);
            string AMPM = "AM";
            if (Time < 1200)
                AMPM = "AM";
            else
                AMPM = "PM";

            if (Time >= 1300) Time = Time - 1200;

            //Make the time 4 digit first
            sTime = "0000" + sTime;
            sTime = sTime.Substring(sTime.Length - 4);
            //Split it with :
            sTime = sTime.Substring(0, 2) + ":" + sTime.Substring(sTime.Length - 2);
            //Add the AMPM
            sTime = sTime + AMPM;

            return sTime;
        }

        public static void WriteEventLog(string Log)
        {
            if (oEventLog != null)
            {
                oEventLog.WriteLine(DateTime.Now + "  " + Log);
            }
        }

        public static int MsgBox(string Message, int DefaultButton, string Caption1, string Caption2, string Caption3)
        {
            int Result = -1;

            System.Threading.Timer oTimer = new System.Threading.Timer(new System.Threading.TimerCallback(KeepWindowsAlive));
            oTimer.Change(0, 60 * 1000);
            try
            {
                Result = SBO_Application.MessageBox(Message, DefaultButton, Caption1, Caption2, Caption3);
            }
            catch
            { }
            finally
            {
                oTimer.Dispose();
            }
            return Result;
        }

        private static void KeepWindowsAlive(object State)
        {
            try
            {
                SBO_Application.RemoveWindowsMessage(BoWindowsMessageType.bo_WM_TIMER, true);
            }
            catch { }
        }

        public static object ExecuteScalar(string SQLQuery)
        {
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRS.DoQuery(SQLQuery);
                return oRS.Fields.Item(0).Value;
            }
            catch
            {
                return null;
            }
            finally
            {
                ReleaseComObject(oRS);
            }
        }

        public static SAPbobsCOM.Recordset ExecuteQuery(string SQLQuery)
        {
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRS.DoQuery(SQLQuery);
                return oRS;
            }
            catch
            {
                return null;
            }
        }

        public static bool ExecuteCommand(string SQLQuery)
        {
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRS.DoQuery(SQLQuery);
                return true;
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Failed executing command. " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                return false;
            }
            finally
            {
                ReleaseComObject(oRS);
            }
        }

        public static void AddReportType(string TypeName, string ReportName, string frmTypeEX, string RptMnuUID, string ReportFileLocation)
        {

            SAPbobsCOM.Recordset oRS = _oCom.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            string RptTypeCode = "";
            string RptLayOutCode = "";
            try
            {
                string sSQL = string.Format("SELECT CODE FROM RTYP WHERE ADD_NAME = '{0}' AND FRM_TYPE = '{1}'", SBOAddon.gcAddOnName, frmTypeEX);
                oRS.DoQuery(sSQL);
                if (oRS.EoF)
                {
                    //Report Type Not Exists. Add a new one
                    SAPbobsCOM.ReportTypesService rptTypeService = (SAPbobsCOM.ReportTypesService)_oCom.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService);
                    SAPbobsCOM.ReportType newType = (SAPbobsCOM.ReportType)rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType);
                    newType.TypeName = TypeName;
                    newType.AddonName = SBOAddon.gcAddOnName;
                    newType.AddonFormType = frmTypeEX;
                    newType.MenuID = RptMnuUID;
                    SAPbobsCOM.ReportTypeParams newTypeParam = rptTypeService.AddReportType(newType);

                    RptTypeCode = newTypeParam.TypeCode;

                }
                else
                    RptTypeCode = oRS.Fields.Item("CODE").Value.ToString();

                sSQL = string.Format("SELECT DocCode FROM RDOC WHERE TypeCode = '{0}' And DocName = '{1}'", RptTypeCode, ReportName);
                oRS.DoQuery(sSQL);
                if (oRS.EoF)
                {
                    //Report Not Exists. Create
                    SAPbobsCOM.ReportLayoutsService rptService = (SAPbobsCOM.ReportLayoutsService)_oCom.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService);
                    SAPbobsCOM.ReportLayout newReport = (SAPbobsCOM.ReportLayout)rptService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayout);
                    newReport.Author = _oCom.UserName;
                    newReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal;
                    newReport.Name = ReportName;
                    newReport.TypeCode = RptTypeCode;
                    SAPbobsCOM.ReportLayoutParams newReportParam = rptService.AddReportLayout(newReport);

                    //Set as Default Report
                    SAPbobsCOM.ReportTypesService rptTypeService = (SAPbobsCOM.ReportTypesService)_oCom.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService);
                    SAPbobsCOM.ReportTypeParams rptTypeParams = rptTypeService.GetDataInterface(ReportTypesServiceDataInterfaces.rtsReportTypeParams) as SAPbobsCOM.ReportTypeParams;
                    rptTypeParams.TypeCode = RptTypeCode;
                    SAPbobsCOM.ReportType rptType = (SAPbobsCOM.ReportType)rptTypeService.GetReportType(rptTypeParams);
                    rptType.DefaultReportLayout = newReportParam.LayoutCode;
                    rptTypeService.UpdateReportType(rptType);
                    RptLayOutCode = newReportParam.LayoutCode;

                }
                else
                {
                    RptLayOutCode = oRS.Fields.Item("DocCode").Value.ToString();
                }


                //Upload the crystal report file to DB using blob
                sSQL = string.Format("SELECT Template FROM RDOC WHERE DocCode = '{0}' AND Template IS NOT NULL", RptLayOutCode);
                if (oRS.EoF)
                {
                    SAPbobsCOM.BlobParams oBlobParams = (SAPbobsCOM.BlobParams)_oCom.GetCompanyService().GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams);
                    oBlobParams.Table = "RDOC";
                    oBlobParams.Field = "Template";
                    SAPbobsCOM.BlobTableKeySegment oKeySegment = oBlobParams.BlobTableKeySegments.Add();
                    oKeySegment.Name = "DocCode";
                    oKeySegment.Value = RptLayOutCode;

                    System.IO.FileStream oFile = new System.IO.FileStream(ReportFileLocation, System.IO.FileMode.Open);
                    int fileSize = (int)oFile.Length;
                    byte[] buf = new byte[fileSize];
                    oFile.Read(buf, 0, fileSize);
                    oFile.Dispose();

                    SAPbobsCOM.Blob oBlob = (SAPbobsCOM.Blob)_oCom.GetCompanyService().GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlob);
                    oBlob.Content = Convert.ToBase64String(buf, 0, fileSize);
                    eCommon.oCompany.GetCompanyService().SetBlob(oBlobParams, oBlob);

                }

            }
            catch (Exception ex)
            {
                eCommon.SBO_Application.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                eCommon.ReleaseComObject(oRS);
            }
        }

        public static void AddReportTypeEx(string TypeName, string ReportName, string frmTypeEX, string RptMnuUID, string ReportResourceName)
        {

            SAPbobsCOM.Recordset oRS = _oCom.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            string RptTypeCode = "";
            string RptLayOutCode = "";
            try
            {
                string sSQL = string.Format("SELECT CODE FROM RTYP WHERE ADD_NAME = '{0}' AND FRM_TYPE = '{1}'", SBOAddon.gcAddOnName, frmTypeEX);
                oRS.DoQuery(sSQL);
                if (oRS.EoF)
                {
                    //Report Type Not Exists. Add a new one
                    SAPbobsCOM.ReportTypesService rptTypeService = (SAPbobsCOM.ReportTypesService)_oCom.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService);
                    SAPbobsCOM.ReportType newType = (SAPbobsCOM.ReportType)rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType);
                    newType.TypeName = TypeName;
                    newType.AddonName = SBOAddon.gcAddOnName;
                    newType.AddonFormType = frmTypeEX;
                    newType.MenuID = RptMnuUID;
                    SAPbobsCOM.ReportTypeParams newTypeParam = rptTypeService.AddReportType(newType);

                    RptTypeCode = newTypeParam.TypeCode;

                }
                else
                    RptTypeCode = oRS.Fields.Item("CODE").Value.ToString();

                sSQL = string.Format("SELECT DocCode FROM RDOC WHERE TypeCode = '{0}' And DocName = '{1}'", RptTypeCode, ReportName);
                oRS.DoQuery(sSQL);
                if (oRS.EoF)
                {
                    //Report Not Exists. Create
                    SAPbobsCOM.ReportLayoutsService rptService = (SAPbobsCOM.ReportLayoutsService)_oCom.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService);
                    SAPbobsCOM.ReportLayout newReport = (SAPbobsCOM.ReportLayout)rptService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayout);
                    newReport.Author = _oCom.UserName;
                    newReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal;
                    newReport.Name = ReportName;
                    newReport.TypeCode = RptTypeCode;
                    SAPbobsCOM.ReportLayoutParams newReportParam = rptService.AddReportLayout(newReport);

                    //Set as Default Report
                    SAPbobsCOM.ReportTypesService rptTypeService = (SAPbobsCOM.ReportTypesService)_oCom.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService);
                    SAPbobsCOM.ReportTypeParams rptTypeParams = rptTypeService.GetDataInterface(ReportTypesServiceDataInterfaces.rtsReportTypeParams) as SAPbobsCOM.ReportTypeParams;
                    rptTypeParams.TypeCode = RptTypeCode;
                    SAPbobsCOM.ReportType rptType = (SAPbobsCOM.ReportType)rptTypeService.GetReportType(rptTypeParams);
                    rptType.DefaultReportLayout = newReportParam.LayoutCode;
                    rptTypeService.UpdateReportType(rptType);
                    RptLayOutCode = newReportParam.LayoutCode;

                }
                else
                {
                    RptLayOutCode = oRS.Fields.Item("DocCode").Value.ToString();
                }


                //Upload the crystal report file to DB using blob
                sSQL = string.Format("SELECT Template FROM RDOC WHERE DocCode = '{0}' AND Template IS NOT NULL", RptLayOutCode);
                if (oRS.EoF)
                {
                    SAPbobsCOM.BlobParams oBlobParams = (SAPbobsCOM.BlobParams)_oCom.GetCompanyService().GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams);
                    oBlobParams.Table = "RDOC";
                    oBlobParams.Field = "Template";
                    SAPbobsCOM.BlobTableKeySegment oKeySegment = oBlobParams.BlobTableKeySegments.Add();
                    oKeySegment.Name = "DocCode";
                    oKeySegment.Value = RptLayOutCode;


                    String ResourceName = string.Format("{0}.Src.Resource.{1}", System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, ReportResourceName);
                    byte[] buf = GetByteResource(ResourceName);
                    SAPbobsCOM.Blob oBlob = (SAPbobsCOM.Blob)_oCom.GetCompanyService().GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlob);
                    oBlob.Content = Convert.ToBase64String(buf, 0, buf.Length);
                    eCommon.oCompany.GetCompanyService().SetBlob(oBlobParams, oBlob);

                }

            }
            catch (Exception ex)
            {
                eCommon.SBO_Application.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                eCommon.ReleaseComObject(oRS);
            }
        }
        public static System.Collections.Hashtable CollectFormsAttribute()
        {
            System.Collections.Hashtable oTable = new System.Collections.Hashtable();
            string NameSpace = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            Type FormAttr = Type.GetType(string.Format("{0}.FormAttribute", NameSpace));

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
                    foreach (System.Attribute Attr in type.GetCustomAttributes(FormAttr, false))
                    {
                        FormAttribute frmAttr = (FormAttribute)Attr;
                        frmAttr.TypeName = type.FullName;
                        if (!oTable.ContainsKey(frmAttr.FormType))
                            oTable.Add(frmAttr.FormType, frmAttr);
                        else
                            SBO_Application.MessageBox(string.Format("The form type {0} can not be registered twice", frmAttr.FormType));
                    }

                }
            }

            return oTable;

        }

        public static System.Collections.Hashtable CollectAuthorizationAttribute()
        {
            System.Collections.Hashtable oTable = new System.Collections.Hashtable();
            string NameSpace = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            Type AuthAttrType = Type.GetType(string.Format("{0}.AuthorizationAttribute", NameSpace));

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
                    foreach (System.Attribute Attr in type.GetCustomAttributes(AuthAttrType, false))
                    {
                        AuthorizationAttribute AuthAttr = (AuthorizationAttribute)Attr;
                        oTable.Add(AuthAttr.FormType, AuthAttr);
                    }
                }
            }

            return oTable;

        }




        public static string Setting_GetValue(String AddOn, String Key)
        {
            String sResult = "";
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            String sSQL = "SELECT U_KeyValue FROM [@TWM_CSTSET] WHERE U_AddOn = '" + AddOn + "' AND U_Key = '" + Key + "'";

            try
            {
                oRS.DoQuery(sSQL);
                if (!oRS.EoF)
                    sResult = oRS.Fields.Item(0).Value.ToString();
                else
                    sResult = null;
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("Failed retrieving setting value. \n\r" + ex.Message);
                sResult = null;
            }


            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;

            GC.Collect();
            return sResult;
        }

        public static bool Setting_Update(String AddOn, String Key, String NewValue)
        {
            bool bResult = false;
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            String sSQL = "UPDATE [@TWM_CSTSET] SET U_KeyValue = '" + NewValue.Replace("'", "''") + "', U_UpdateOn = '" + DateTime.Now.ToString("yyyyMMdd hhmmss") + "', U_UpdateBy = '" + oCompany.UserName + "' WHERE U_AddOn = '" + AddOn + "' AND U_Key = '" + Key.Replace("'", "''") + "'";

            try
            {
                oRS.DoQuery(sSQL);
                bResult = true;
            }
            catch (Exception ex)
            {
                oEventLog.WriteLine(DateTime.Now + Filler + "Failed updating setting. " + sSQL + ". " + ex.Message);
                bResult = false;
            }


            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;

            GC.Collect();
            return bResult;

        }

        public static bool Setting_Insert(String AddOn, String Key, String NewValue, String Description)
        {
            bool bResult = false;
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            String sCode = "";
            oCompany.StartTransaction();
            oRS.DoQuery("Select isnull(Max(Convert(Int, Code)),0) From [@TWM_CSTSET]");
            sCode = ((int)oRS.Fields.Item(0).Value + 1).ToString();
            String sSQL = "INSERT INTO [@TWM_CSTSET] (Code, Name, U_AddOn, U_Key, U_Dcrption, U_KeyValue) Values ('" + sCode + "','" + sCode + "', '" + AddOn + "', '" + Key.Replace("'", "''") + "', '" + Description.Replace("'", "''") + "', '" + NewValue.Replace("'", "''") + "')";

            try
            {
                oRS.DoQuery(sSQL);
                bResult = true;
            }
            catch (Exception ex)
            {
                oEventLog.WriteLine(DateTime.Now + Filler + "Failed inserting setting. " + sSQL + ". " + ex.Message);
                bResult = false;
            }
            finally
            {
                if (oCompany.InTransaction) oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }


            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;

            GC.Collect();
            return bResult;
        }

        public static String GetXMLResource(String ResourceName)
        {
            String sContent = "";
            System.IO.Stream oStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName);
            if (oStream != null)
            {
                using (System.IO.StreamReader oReader = new System.IO.StreamReader(oStream))
                {
                    sContent = oReader.ReadToEnd();
                }
            }

            return sContent;
        }

        public static Byte[] GetByteResource(String ResourceName)
        {
            
            System.IO.Stream oStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName);
            byte[] buffer = new Byte[oStream.Length];
            if (oStream != null)
            {
                oStream.Read(buffer, 0, (int)oStream.Length);
            }

            return buffer;
        }

        public static void Combo_FillValidValues(SAPbouiCOM.ComboBox oCBO, String sSQL, bool Clear = true)
        {
            if (Clear)
            {
                int iCBOValid = oCBO.ValidValues.Count;
                for (int i = 0; i < iCBOValid; i++)
                {
                    oCBO.ValidValues.Remove(0, BoSearchKey.psk_Index);
                }

                SAPbobsCOM.Recordset oRS = ExecuteQuery(sSQL);
                for (int i = 0; i < oRS.RecordCount; i++)
                {
                    oCBO.ValidValues.Add(oRS.Fields.Item(0).Value.ToString().Trim(), oRS.Fields.Item(1).Value.ToString().Trim());
                    oRS.MoveNext();
                }
            }
        }

        public static void Combo_FillValidValues(SAPbouiCOM.ComboBoxColumn oCBO, String sSQL, bool Clear = true)
        {
            if (Clear)
            {
                int iCBOValid = oCBO.ValidValues.Count;
                for (int i = 0; i < iCBOValid; i++)
                {
                    oCBO.ValidValues.Remove(0, BoSearchKey.psk_Index);
                }

                SAPbobsCOM.Recordset oRS = ExecuteQuery(sSQL);
                for (int i = 0; i < oRS.RecordCount; i++)
                {
                    oCBO.ValidValues.Add(oRS.Fields.Item(0).Value.ToString().Trim(), oRS.Fields.Item(1).Value.ToString().Trim());
                    oRS.MoveNext();
                }
            }
        }

        public static void Combo_FillValidValues(SAPbouiCOM.Column oCBO, String sSQL, bool Clear = true)
        {
            if (Clear)
            {
                int iCBOValid = oCBO.ValidValues.Count;
                for (int i = 0; i < iCBOValid; i++)
                {
                    oCBO.ValidValues.Remove(0, BoSearchKey.psk_Index);
                }

                SAPbobsCOM.Recordset oRS = ExecuteQuery(sSQL);
                for (int i = 0; i < oRS.RecordCount; i++)
                {
                    oCBO.ValidValues.Add(oRS.Fields.Item(0).Value.ToString().Trim(), oRS.Fields.Item(1).Value.ToString().Trim());
                    oRS.MoveNext();
                }
            }
        }

        public static void Combo_FillValidValues(SAPbouiCOM.ComboBox oCBO, String TableName, String FieldName, bool Clear = true)
        {
            if (Clear)
            {
                int iCBOValid = oCBO.ValidValues.Count;
                for (int i = 0; i < iCBOValid; i++)
                {
                    oCBO.ValidValues.Remove(0, BoSearchKey.psk_Index);
                }

                string sSQL = string.Format("SELECT T1.FldValue, T1.Descr FROM CUFD T0 JOIN UFD1 T1 ON T0.TableID = T1.TableID AND T0.FieldID = T1.FieldID WHERE T0.TableID = '{0}' AND T0.AliasID = '{1}'", TableName, FieldName);
                SAPbobsCOM.Recordset oRS = ExecuteQuery(sSQL);
                for (int i = 0; i < oRS.RecordCount; i++)
                {
                    oCBO.ValidValues.Add(oRS.Fields.Item(0).Value.ToString().Trim(), oRS.Fields.Item(1).Value.ToString().Trim());
                    oRS.MoveNext();
                }
            }
        }

        public static void Combo_FillValidValues(SAPbouiCOM.ComboBoxColumn oCBO, String TableName, String FieldName, bool Clear = true)
        {
            if (Clear)
            {
                int iCBOValid = oCBO.ValidValues.Count;
                for (int i = 0; i < iCBOValid; i++)
                {
                    oCBO.ValidValues.Remove(0, BoSearchKey.psk_Index);
                }

                string sSQL = string.Format("SELECT T1.FldValue, T1.Descr FROM CUFD T0 JOIN UFD1 T1 ON T0.TableID = T1.TableID AND T0.FieldID = T1.FieldID WHERE T0.TableID = '{0}' AND T0.AliasID = '{1}'", TableName, FieldName);
                SAPbobsCOM.Recordset oRS = ExecuteQuery(sSQL);
                for (int i = 0; i < oRS.RecordCount; i++)
                {
                    oCBO.ValidValues.Add(oRS.Fields.Item(0).Value.ToString().Trim(), oRS.Fields.Item(1).Value.ToString().Trim());
                    oRS.MoveNext();
                }
            }
        }

        public static void Combo_FillValidValues(SAPbouiCOM.Column oCBO, String TableName, String FieldName, bool Clear = true)
        {
            if (Clear)
            {
                int iCBOValid = oCBO.ValidValues.Count;
                for (int i = 0; i < iCBOValid; i++)
                {
                    oCBO.ValidValues.Remove(0, BoSearchKey.psk_Index);
                }

                string sSQL = string.Format("SELECT T1.FldValue, T1.Descr FROM CUFD T0 JOIN UFD1 T1 ON T0.TableID = T1.TableID AND T0.FieldID = T1.FieldID WHERE T0.TableID = '{0}' AND T0.AliasID = '{1}'", TableName, FieldName);
                SAPbobsCOM.Recordset oRS = ExecuteQuery(sSQL);
                for (int i = 0; i < oRS.RecordCount; i++)
                {
                    oCBO.ValidValues.Add(oRS.Fields.Item(0).Value.ToString().Trim(), oRS.Fields.Item(1).Value.ToString().Trim());
                    oRS.MoveNext();
                }
            }
        }

        public static string GetFormattedBPAddress(String CardCode, String AdresType, String AddressName)
        {
            String sSQL = @"
SELECT T3.""Format"", T1.""Address"", T1.Street	""$1"", T1.City ""$2"", T1.ZipCode	""$3"", T1.County	""$4"", T1.""State"" ""$5"", T1.Country	""$6"", T1.Block	""$7"", T1.Building	""$B"", T1.AdresType	""$A"", T1.StreetNo	""$T"", T1.Address2	""$8"", T1.Address3	""$9"", T2.Name ""$D""
FROM CRD1 T1 	JOIN OCRY T2 ON T1.Country = T2.Code
	JOIN OADF T3 ON T2.AddrFormat = T3.Code
WHERE T1.CardCode = '{0}'
	AND T1.AdresType = '{1}'
	AND Address Like '{2}'";
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                oRS = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                sSQL = String.Format(sSQL, CardCode, AdresType, AddressName);
                oRS.DoQuery(sSQL);
                String sFormat = oRS.Fields.Item("Format").Value.ToString();
                //Split $N into muliple Lines first
                string[] AddressLines = sFormat.Split(new String[] {"$N"}, StringSplitOptions.None);
                string sFinalAddress = "";
                foreach (string Line in AddressLines)
                {
                    System.Collections.Generic.List<String> oLineItems = new System.Collections.Generic.List<string>();
                    int iPosFrom = 0;
                    String sFinalLine = "";
                    while (iPosFrom < Line.Length)
                    {
                        int iPosTo = Line.IndexOf('$', iPosFrom+1)+1;
                        String sItem = "";
                        if (iPosTo > 0)
                        {
                            sItem = Line.Substring(iPosFrom, iPosTo-iPosFrom+1);
                            iPosFrom = iPosTo + 1;
                        }
                        else
                        {
                            sItem = Line.Substring(iPosFrom);
                            iPosFrom = Line.Length;
                        }
                        
                        bool isDescription = false; //D
                        bool isUpperCase = false;   //U
                        bool isLowerCase = false;   //O
                        bool isCapitalised = false; //T
                        if (sItem.EndsWith("D]")) isDescription = true;
                        if (sItem.Contains("[U")) isUpperCase = true;
                        if (sItem.Contains("[O")) isLowerCase = true;
                        if (sItem.Contains("[T")) isCapitalised = true;
                        int iBrackStart = sItem.IndexOf('[');
                        if (iBrackStart > -1)
                        {
                            int iBrackEnd = sItem.IndexOf(']') +1;
                            sItem = sItem.Remove(iBrackStart, iBrackEnd - iBrackStart);
                        }
                        sItem = sItem.Replace("$L", "");
                        if (isUpperCase)
                        {
                            sItem = sItem.Replace("$1", oRS.Fields.Item("$1").Value.ToString().Trim().ToUpper());
                            sItem = sItem.Replace("$2", oRS.Fields.Item("$2").Value.ToString().Trim().ToUpper());
                            sItem = sItem.Replace("$3", oRS.Fields.Item("$3").Value.ToString().Trim().ToUpper());
                            sItem = sItem.Replace("$4", oRS.Fields.Item("$4").Value.ToString().Trim().ToUpper());
                            sItem = sItem.Replace("$5", oRS.Fields.Item("$5").Value.ToString().Trim().ToUpper());
                            if (isDescription)
                                sItem = sItem.Replace("$6", oRS.Fields.Item("$D").Value.ToString().Trim().ToUpper());
                            else
                                sItem = sItem.Replace("$6", oRS.Fields.Item("$6").Value.ToString().Trim().ToUpper());
                            sItem = sItem.Replace("$7", oRS.Fields.Item("$7").Value.ToString().Trim().ToUpper());
                            sItem = sItem.Replace("$B", oRS.Fields.Item("$B").Value.ToString().Trim().ToUpper());
                            sItem = sItem.Replace("$A", oRS.Fields.Item("$A").Value.ToString().Trim().ToUpper());
                            sItem = sItem.Replace("$T", oRS.Fields.Item("$T").Value.ToString().Trim().ToUpper());
                            sItem = sItem.Replace("$8", oRS.Fields.Item("$8").Value.ToString().Trim().ToUpper());
                            sItem = sItem.Replace("$9", oRS.Fields.Item("$9").Value.ToString().Trim().ToUpper());
                        }
                        else if (isLowerCase)
                        {
                            sItem = sItem.Replace("$1", oRS.Fields.Item("$1").Value.ToString().Trim().ToLower());
                            sItem = sItem.Replace("$2", oRS.Fields.Item("$2").Value.ToString().Trim().ToLower());
                            sItem = sItem.Replace("$3", oRS.Fields.Item("$3").Value.ToString().Trim().ToLower());
                            sItem = sItem.Replace("$4", oRS.Fields.Item("$4").Value.ToString().Trim().ToLower());
                            sItem = sItem.Replace("$5", oRS.Fields.Item("$5").Value.ToString().Trim().ToLower());
                            if (isDescription)
                                sItem = sItem.Replace("$6", oRS.Fields.Item("$D").Value.ToString().Trim().ToLower());
                            else
                                sItem = sItem.Replace("$6", oRS.Fields.Item("$6").Value.ToString().Trim().ToLower());
                            sItem = sItem.Replace("$7", oRS.Fields.Item("$7").Value.ToString().Trim().ToLower());
                            sItem = sItem.Replace("$B", oRS.Fields.Item("$B").Value.ToString().Trim().ToLower());
                            sItem = sItem.Replace("$A", oRS.Fields.Item("$A").Value.ToString().Trim().ToLower());
                            sItem = sItem.Replace("$T", oRS.Fields.Item("$T").Value.ToString().Trim().ToLower());
                            sItem = sItem.Replace("$8", oRS.Fields.Item("$8").Value.ToString().Trim().ToLower());
                            sItem = sItem.Replace("$9", oRS.Fields.Item("$9").Value.ToString().Trim().ToLower());
                        }
                        else if (isCapitalised)
                        {
                            sItem = sItem.Replace("$1", oRS.Fields.Item("$1").Value.ToString().Trim().ToUpperInvariant());
                            sItem = sItem.Replace("$2", oRS.Fields.Item("$2").Value.ToString().Trim().ToUpperInvariant());
                            sItem = sItem.Replace("$3", oRS.Fields.Item("$3").Value.ToString().Trim().ToUpperInvariant());
                            sItem = sItem.Replace("$4", oRS.Fields.Item("$4").Value.ToString().Trim().ToUpperInvariant());
                            sItem = sItem.Replace("$5", oRS.Fields.Item("$5").Value.ToString().Trim().ToUpperInvariant());
                            if (isDescription)
                                sItem = sItem.Replace("$6", oRS.Fields.Item("$D").Value.ToString().Trim().ToUpperInvariant());
                            else
                                sItem = sItem.Replace("$6", oRS.Fields.Item("$6").Value.ToString().Trim().ToUpperInvariant());
                            sItem = sItem.Replace("$7", oRS.Fields.Item("$7").Value.ToString().Trim().ToUpperInvariant());
                            sItem = sItem.Replace("$B", oRS.Fields.Item("$B").Value.ToString().Trim().ToUpperInvariant());
                            sItem = sItem.Replace("$A", oRS.Fields.Item("$A").Value.ToString().Trim().ToUpperInvariant());
                            sItem = sItem.Replace("$T", oRS.Fields.Item("$T").Value.ToString().Trim().ToUpperInvariant());
                            sItem = sItem.Replace("$8", oRS.Fields.Item("$8").Value.ToString().Trim().ToUpperInvariant());
                            sItem = sItem.Replace("$9", oRS.Fields.Item("$9").Value.ToString().Trim().ToUpperInvariant());
                        }
                        else
                        {
                            sItem = sItem.Replace("$1", oRS.Fields.Item("$1").Value.ToString().Trim());
                            sItem = sItem.Replace("$2", oRS.Fields.Item("$2").Value.ToString().Trim());
                            sItem = sItem.Replace("$3", oRS.Fields.Item("$3").Value.ToString().Trim());
                            sItem = sItem.Replace("$4", oRS.Fields.Item("$4").Value.ToString().Trim());
                            sItem = sItem.Replace("$5", oRS.Fields.Item("$5").Value.ToString().Trim());
                            if (isDescription)
                                sItem = sItem.Replace("$6", oRS.Fields.Item("$D").Value.ToString().Trim());
                            else
                                sItem = sItem.Replace("$6", oRS.Fields.Item("$6").Value.ToString().Trim());
                            sItem = sItem.Replace("$7", oRS.Fields.Item("$7").Value.ToString().Trim());
                            sItem = sItem.Replace("$B", oRS.Fields.Item("$B").Value.ToString().Trim());
                            sItem = sItem.Replace("$A", oRS.Fields.Item("$A").Value.ToString().Trim());
                            sItem = sItem.Replace("$T", oRS.Fields.Item("$T").Value.ToString().Trim());
                            sItem = sItem.Replace("$8", oRS.Fields.Item("$8").Value.ToString().Trim());
                            sItem = sItem.Replace("$9", oRS.Fields.Item("$9").Value.ToString().Trim());

                        }

                        sFinalLine += sItem;
                    }
                    sFinalAddress += sFinalLine + "\r\n";
                }

                return sFinalAddress;
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
            finally
            {
                ReleaseComObject(oRS);
            }
        }

        public static double GetLineTotal(String CardCode, String ItemCode, double Quantity, DateTime DocDate)
        {
            SAPbobsCOM.CompanyService oCS = eCommon.oCompany.GetCompanyService();
            SAPbobsCOM.ItemPriceParams oParams = oCS.GetDataInterface(CompanyServiceDataInterfaces.csdiItemPriceParams) as SAPbobsCOM.ItemPriceParams;
            SAPbobsCOM.ItemPriceReturnParams oRet = null;
            oParams.CardCode = CardCode;
            oParams.Date = DocDate;
            oParams.ItemCode = ItemCode;
            oParams.InventoryQuantity = Quantity;

            oRet = oCS.GetItemPrice(oParams);

            return oRet.Price;

        }

        public static double GetItemPrice(String CardCode, String ItemCode, double Quantity, DateTime DocDate)
        {
            SAPbobsCOM.SBObob sBOB = null;
            SAPbobsCOM.Recordset oRS = null;
            Double dPrice = 0;

            sBOB = eCommon.oCompany.GetBusinessObject(BoObjectTypes.BoBridge) as SAPbobsCOM.SBObob;
            oRS = sBOB.GetItemPrice(CardCode, ItemCode, Quantity, DocDate);
            dPrice = (double)oRS.Fields.Item(0).Value;


            return dPrice;

        }

        public static void UDO_SetGeneralData(ref SAPbobsCOM.GeneralData oGD,ref SAPbouiCOM.DBDataSource oDB, int iRow, String FieldName)
        {
            switch (oDB.Fields.Item(FieldName).Type)
            {
                case SAPbouiCOM.BoFieldsType.ft_Date:
                    if (oDB.GetValue(oDB.Fields.Item(FieldName).Name, iRow) != "")
                        oGD.SetProperty(oDB.Fields.Item(FieldName).Name, DateTime.ParseExact(oDB.GetValue(oDB.Fields.Item(FieldName).Name, iRow), "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                    else
                        oGD.SetProperty(oDB.Fields.Item(FieldName).Name, "");
                    break;
                case SAPbouiCOM.BoFieldsType.ft_Float:
                case SAPbouiCOM.BoFieldsType.ft_Measure:
                case SAPbouiCOM.BoFieldsType.ft_Percent:
                case SAPbouiCOM.BoFieldsType.ft_Price:
                case SAPbouiCOM.BoFieldsType.ft_Quantity:
                case SAPbouiCOM.BoFieldsType.ft_Rate:
                case SAPbouiCOM.BoFieldsType.ft_Sum:
                    if (oDB.GetValue(oDB.Fields.Item(FieldName).Name, iRow) != "")
                        oGD.SetProperty(oDB.Fields.Item(FieldName).Name, double.Parse(oDB.GetValue(oDB.Fields.Item(FieldName).Name, iRow)));
                    else
                        oGD.SetProperty(oDB.Fields.Item(FieldName).Name, iRow);

                    break;
                case SAPbouiCOM.BoFieldsType.ft_Integer:
                case SAPbouiCOM.BoFieldsType.ft_ShortNumber:
                    if (oDB.GetValue(oDB.Fields.Item(FieldName).Name, iRow) != "")
                        oGD.SetProperty(oDB.Fields.Item(FieldName).Name, int.Parse(oDB.GetValue(oDB.Fields.Item(FieldName).Name, iRow)));
                    else
                        oGD.SetProperty(oDB.Fields.Item(FieldName).Name, "");
                    break;
                default:
                    oGD.SetProperty(oDB.Fields.Item(FieldName).Name, oDB.GetValue(oDB.Fields.Item(FieldName).Name, iRow).Trim());
                    break;
            }
        }

        public static String GetParentFormUID(SAPbouiCOM.Form oForm)
        {
            String ParentFormUID = "";
            for (int i = 0; i < SBO_Application.Forms.Count; i++)
            {
                if (SBO_Application.Forms.Item(i).UDFFormUID == oForm.UniqueID)
                {
                    ParentFormUID = SBO_Application.Forms.Item(i).UniqueID;
                    break;
                }
            }

            return ParentFormUID;
        }

        public static System.Data.DataTable DataTable_Get_Net_DataTable(SAPbouiCOM.DataTable dtInput)
        {

            System.Xml.XmlDocument oXML = new System.Xml.XmlDocument();
            oXML.LoadXml(dtInput.SerializeAsXML(BoDataTableXmlSelect.dxs_All) );

            System.Data.DataTable dtResult = new System.Data.DataTable();
            System.Xml.XmlNodeList oList = oXML.GetElementsByTagName("DataTable");
            String TableUID = "";
            if (oList.Count > 0)
            {
                try
                {
                    TableUID = oList[0].Attributes["Uid"].Value;
                    dtResult.TableName = TableUID;
                }
                catch { }
            }

            //Columns
            oList = oXML.GetElementsByTagName("Column");
            foreach (System.Xml.XmlNode oColumn in oList)
            {
                String Name = oColumn.Attributes["Uid"].Value;
                switch (oColumn.Attributes["Type"].Value)
                {
                    case "0":   //Undefined"
                    case "1":   //String
                    case "3":   //Text
                        dtResult.Columns.Add(Name, typeof(String));
                        break;
                    case "2":   //Integer
                    case "6":   //Short Number
                        dtResult.Columns.Add(Name, typeof(Int32));
                        break;
                    case "4":   //Date
                        dtResult.Columns.Add(Name, typeof(DateTime));
                        break;
                    case "5":   //Double
                    case "7":   //Quantity
                    case "8":   //Price
                    case "9":   //Rate
                    case "10":  //Measure
                    case "11":  //Sum
                    case "12":  //Percent
                        dtResult.Columns.Add(Name, typeof(Double));
                        break;
                }
            }

            //Add the rows
            oList = oXML.GetElementsByTagName("Row");
            foreach (System.Xml.XmlNode oRow in oList)
            {
                dtResult.Rows.Add();
                System.Xml.XmlNodeList oCells = oRow.ChildNodes[0].ChildNodes;
                foreach (System.Xml.XmlNode oCell in oCells)
                {
                    String ColumnName = oCell.ChildNodes[0].InnerText;
                    String Value = oCell.ChildNodes[1].InnerText;
                    switch (dtResult.Columns[ColumnName].DataType.UnderlyingSystemType.ToString())
                    {
                        case "System.String":
                            dtResult.Rows[dtResult.Rows.Count - 1][ColumnName] = Value;
                            break;
                        case "System.Int32":
                            if (Value != "")
                                dtResult.Rows[dtResult.Rows.Count - 1][ColumnName] = Int32.Parse(Value);
                            break;
                        case "System.DateTime":
                            if (Value == "00000000")
                            {
                                dtResult.Rows[dtResult.Rows.Count - 1][ColumnName] = DateTime.FromOADate(0);
                            }
                            else if (Value == "")
                            {
                                dtResult.Rows[dtResult.Rows.Count - 1][ColumnName] = DateTime.FromOADate(0);
                            }
                            else
                            {
                                dtResult.Rows[dtResult.Rows.Count - 1][ColumnName] = DateTime.ParseExact(Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo);
                            }
                            break;
                        case "System.Double":
                            if (Value != "")
                                dtResult.Rows[dtResult.Rows.Count - 1][ColumnName] = Double.Parse(Value);
                            break;
                    }
                }
            }

            return dtResult;
        }

    } //End Class
}   //End NameSpace





