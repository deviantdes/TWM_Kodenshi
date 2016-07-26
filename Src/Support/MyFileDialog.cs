using System;
using System.Collections.Generic;
using System.Text;
using System.Management;
using System.Windows.Forms;

namespace SBOCustom
{
    public class WindowWrapper : System.Windows.Forms.IWin32Window
    {
        public WindowWrapper(IntPtr handle)
        {
            _hwnd = handle;
        }

        public IntPtr Handle
        {
            get { return _hwnd; }
        }

        private IntPtr _hwnd;
    }

    public class MyFileDialog:IDisposable
    {
        private SAPbouiCOM.Application _App;
        private  string FolderName;
        private  string FileName;
        private  string[] FileNames;
        private  int cParentID;
        private  string cMachineName;
        private bool MultiFile = false;

        public  string cDlgDefaultDir = "";
        public  string cDlgDefaultExt = "";
        public  string cDlgTitle = "Browse Folder";

        public MyFileDialog(SAPbouiCOM.Application oApp)
        {
            _App = oApp;
        }

        public  String OpenFileDialog()
        {
            MultiFile = false;
            String sFile = FindFile();
            return sFile;
        }

        public  string[] OpenMultiFileDialog()
        {
            MultiFile = true;
            String[] sFiles = FindFiles();
            return sFiles;
        }

        public string OpenFolderDialog()
        {
            String sFolder = FindFolder();
            return sFolder;
        }

        private  string FindFile()
        {
            try
            {
                System.Diagnostics.Process cProcess = System.Diagnostics.Process.GetCurrentProcess();
                cParentID = getProcessParentID(cProcess.ProcessName, cProcess.Id);
                System.Diagnostics.Process cParent = System.Diagnostics.Process.GetProcessById(cParentID, cProcess.MachineName);
                cMachineName = cProcess.MachineName;
                if (cProcess.ProcessName.EndsWith("vshost")) // 'This part is for RUNTIME Mode
                {
                    _App.StatusBar.SetText(cProcess.Id + " 2N " + cProcess.ProcessName + " << " + cParent.Id + " N " + cParent.ProcessName + " <<< This is DESIGN-TIME >>>", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    cParentID = 0;
                }
            }
            catch
            {
                _App.StatusBar.SetText("Process Reading Fail!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            System.Threading.Thread ShowFileBrowserThread;
            System.Threading.Timer oTimer = new System.Threading.Timer(new System.Threading.TimerCallback(TimerKeepAlive));
            oTimer.Change(0, 60 * 1000);    //Set the timer for 
            try
            {
                ShowFileBrowserThread = new System.Threading.Thread(new System.Threading.ThreadStart(ShowFileBrowser));
                if (ShowFileBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
                {
                    ShowFileBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA);
                    ShowFileBrowserThread.Start();
                }
                else if (ShowFileBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
                {
                    ShowFileBrowserThread.Start();
                    ShowFileBrowserThread.Join();
                }

                while (ShowFileBrowserThread.ThreadState == System.Threading.ThreadState.Running)
                {
                    System.Windows.Forms.Application.DoEvents();
                }


                if (FileName != "")
                    return FileName;

            }
            catch (Exception ex)
            {
                _App.MessageBox("FindFile. " + ex.Message, 1, "OK", null, null);
            }
            finally
            {
                oTimer.Dispose();
            }
            return "";
        }

        private  string[] FindFiles()
        {
            try
            {
                System.Diagnostics.Process cProcess = System.Diagnostics.Process.GetCurrentProcess();
                cParentID = getProcessParentID(cProcess.ProcessName, cProcess.Id);
                System.Diagnostics.Process cParent = System.Diagnostics.Process.GetProcessById(cParentID, cProcess.MachineName);
                cMachineName = cProcess.MachineName;
                if (cProcess.ProcessName.EndsWith("vshost")) // 'This part is for RUNTIME Mode
                {
                    _App.StatusBar.SetText(cProcess.Id + " 2N " + cProcess.ProcessName + " << " + cParent.Id + " N " + cParent.ProcessName + " <<< This is DESIGN-TIME >>>", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    cParentID = 0;
                }
            }
            catch
            {
                _App.StatusBar.SetText("Process Reading Fail!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            System.Threading.Thread ShowFileBrowserThread;
            System.Threading.Timer oTimer = new System.Threading.Timer(new System.Threading.TimerCallback(TimerKeepAlive));
            oTimer.Change(0, 60 * 1000);    //Set the timer for 
            try
            {
                ShowFileBrowserThread = new System.Threading.Thread(new System.Threading.ThreadStart(ShowFileBrowser));
                if (ShowFileBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
                {
                    ShowFileBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA);
                    ShowFileBrowserThread.Start();
                }
                else if (ShowFileBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
                {
                    ShowFileBrowserThread.Start();
                    ShowFileBrowserThread.Join();
                }

                while (ShowFileBrowserThread.ThreadState == System.Threading.ThreadState.Running)
                {
                    System.Windows.Forms.Application.DoEvents();
                }



                return FileNames;

            }
            catch (Exception ex)
            {
                _App.MessageBox("FindFile. " + ex.Message, 1, "OK", null, null);
            }
            finally
            {
                oTimer.Dispose();
            }
            return null;
        }
        private  String FindFolder()
        {
            try
            {
                System.Diagnostics.Process cProcess = System.Diagnostics.Process.GetCurrentProcess();
                cParentID = getProcessParentID(cProcess.ProcessName, cProcess.Id);
                System.Diagnostics.Process cParent = System.Diagnostics.Process.GetProcessById(cParentID, cProcess.MachineName);
                cMachineName = cProcess.MachineName;
                if (cProcess.ProcessName.EndsWith("vshost")) // 'This part is for RUNTIME Mode
                {
                    _App.StatusBar.SetText(cProcess.Id + " 2N " + cProcess.ProcessName + " << " + cParent.Id + " N " + cParent.ProcessName + " <<< This is DESIGN-TIME >>>", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    cParentID = 0;
                }
            }
            catch
            {
                _App.StatusBar.SetText("Process Reading Fail!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            System.Threading.Thread ShowFolderBrowserThread;
            System.Threading.Timer oTimer = new System.Threading.Timer(new System.Threading.TimerCallback(TimerKeepAlive));
            oTimer.Change(0, 60 * 1000);

            try
            {
                ShowFolderBrowserThread = new System.Threading.Thread(new System.Threading.ThreadStart(ShowFolderBrowser));
                if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
                {
                    ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA);
                    ShowFolderBrowserThread.Start();
                }
                else if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
                {
                    ShowFolderBrowserThread.Start();
                    ShowFolderBrowserThread.Join();
                }

                while (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Running)
                {
                    System.Windows.Forms.Application.DoEvents();
                }

                if (FolderName != "") return FolderName;

            }
            catch (Exception ex)
            {
                _App.MessageBox("FindFile" + ex.Message, 1, "OK", null, null);
            }
            finally
            {
                oTimer.Dispose();
            }

            return "";
        }

        private  void ShowFileBrowser()
        {
            FileName = "";
            OpenFileDialog FileBrowser = new OpenFileDialog();
            try
            {
                if (cDlgDefaultDir != "")
                {
                    FileBrowser.InitialDirectory = cDlgDefaultDir;
                }
                else
                {
                    FileBrowser.InitialDirectory = Environment.SpecialFolder.MyComputer.ToString();
                }

                if (cDlgTitle != "") FileBrowser.Title = cDlgTitle;
                if (cDlgDefaultExt != "") FileBrowser.Filter = cDlgDefaultExt;
                FileBrowser.Multiselect = MultiFile;

                if (cParentID == 0)
                {
                    System.Diagnostics.Process[] MyProcs = System.Diagnostics.Process.GetProcessesByName("SAP Business One");
                    WindowWrapper MyWindow = new WindowWrapper(MyProcs[0].MainWindowHandle);
                    DialogResult ret = FileBrowser.ShowDialog(MyWindow);

                    if (ret == DialogResult.OK)
                    {
                        if (MultiFile)
                        {
                            FileNames = FileBrowser.FileNames;
                            FileBrowser.Dispose();
                        }
                        else
                        {
                            FileName = FileBrowser.FileName;
                            FileBrowser.Dispose();
                        }
                    }
                    else
                    {
                        System.Windows.Forms.Application.ExitThread();
                    }

                }
                else
                {
                    System.Diagnostics.Process kProcs = System.Diagnostics.Process.GetProcessById(cParentID, cMachineName);
                    WindowWrapper xWindow = new WindowWrapper(kProcs.MainWindowHandle);
                    DialogResult ret = FileBrowser.ShowDialog(xWindow);

                    if (ret == DialogResult.OK)
                    {
                        if (MultiFile)
                        {
                            FileNames = FileBrowser.FileNames;
                            FileBrowser.Dispose();
                        }
                        else
                        {
                            FileName = FileBrowser.FileName;
                            FileBrowser.Dispose();
                        }
                    }
                    else
                    {
                        System.Windows.Forms.Application.ExitThread();
                    }
                }
            }
            catch (Exception ex)
            {
                FileName = "Error " + ex.Message.ToString();
            }
            finally
            {
                FileBrowser.Dispose();
            }



        }

        private  void ShowFolderBrowser()
        {
            FolderName = "";
            FolderBrowserDialog OpenFolder = new FolderBrowserDialog();
            try
            {

                if (cDlgDefaultDir != "") OpenFolder.RootFolder = Environment.SpecialFolder.MyComputer;
                OpenFolder.SelectedPath = cDlgDefaultDir;
                if (cDlgTitle != "") OpenFolder.Description = cDlgTitle;

                if (cParentID == 0)
                {
                    System.Diagnostics.Process[] MyProcs = System.Diagnostics.Process.GetProcessesByName("SAP Business One");
                    WindowWrapper MyWindow = new WindowWrapper(MyProcs[0].MainWindowHandle);
                    DialogResult ret = OpenFolder.ShowDialog(MyWindow);

                    if (ret == DialogResult.OK)
                    {
                        FolderName = OpenFolder.SelectedPath;
                        OpenFolder.Dispose();
                    }
                    else
                    {
                        System.Windows.Forms.Application.ExitThread();
                    }
                }
                else
                {
                    System.Diagnostics.Process kProcs = System.Diagnostics.Process.GetProcessById(cParentID, cMachineName);
                    WindowWrapper xWindow = new WindowWrapper(kProcs.MainWindowHandle);
                    DialogResult ret = OpenFolder.ShowDialog(xWindow);

                    if (ret == DialogResult.OK)
                    {
                        FolderName = OpenFolder.SelectedPath;
                        OpenFolder.Dispose();
                    }
                    else
                    {
                        System.Windows.Forms.Application.ExitThread();
                    }
                }
            }
            catch (Exception ex)
            {
                FolderName = "Error " + ex.Message.ToString();
            }
            finally
            {
                OpenFolder.Dispose();
            }
        }


        //Need to Imports System.Management ( Add Reference )
        private  int getProcessParentID(string cName, int cID)
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

        private  void TimerKeepAlive(object State)
        {
            try
            {
                _App.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, true);
            }
            catch { }
        }

        #region IDisposable Members

        public void Dispose()
        {
            _App = null;
        }

        #endregion
    }
}
