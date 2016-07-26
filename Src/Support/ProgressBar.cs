using System;
using System.Collections.Generic;
using System.Text;

namespace SBOCustom
{
    public class SBOCustom_ProgressBar : IDisposable
    {
        System.Threading.Timer oTimer = null;
        private SAPbouiCOM.Application _App;
        private string _Title = "Progress Bar";
        private int _MaxValue = 100;
        private int _Value = 0;
        private string _ParentFormUID = "";
        private int _X = 0;
        private int _Y = 0;
        private int _BarWidth = 0;
        private SAPbouiCOM.Form _oForm;
        private ProgressType _Type = ProgressType.ValueCount;
        private string _FormUID = "";

        public enum ProgressType
        {
            ValueCount,
            ProgressBar,
            Both
        }

        public SBOCustom_ProgressBar(SAPbouiCOM.Application SBO_App, SAPbouiCOM.Form ParentForm, string Title, int MaxValue, ProgressType Type)
        {
            _App = SBO_App;
            _Title = Title;
            _MaxValue = MaxValue;
            if (ParentForm != null)
            {
                _ParentFormUID = ParentForm.UniqueID;
                _X = ParentForm.Left + Convert.ToInt32(ParentForm.Width / 2) - 125;
                _Y = ParentForm.Top + Convert.ToInt32(ParentForm.Height / 2) - 20;
            }
            else
            {
                _X = SBO_App.Desktop.Width - 250;
                _Y = 100;
            }


            _Type = Type;


            //Create the form
            _CreateProgressBar();
        }

        public int Value
        {
            get
            {
                return _Value;
            }
            set
            {
                _Value = value;
                //Refresh the form
                _oForm = _App.Forms.Item(_FormUID);

                if (_Type == ProgressType.ValueCount)
                {
                    _oForm.Freeze(true);
                    _oForm.DataSources.UserDataSources.Item("txtCount").Value = _Value.ToString();
                    _oForm.Freeze(false);
                }
                else if (_Type == ProgressType.ProgressBar)
                {
                    _oForm.Items.Item("lblBar").Width = Convert.ToInt32((Convert.ToDouble(value) / Convert.ToDouble(_MaxValue)) * _BarWidth);
                }
                else
                {
                    _oForm.Freeze(true);
                    _oForm.DataSources.UserDataSources.Item("txtCount").Value = _Value.ToString();
                    _oForm.Items.Item("lblBar").Width = Convert.ToInt32((Convert.ToDouble(value) / Convert.ToDouble(_MaxValue)) * _BarWidth);
                    _oForm.Freeze(false);
                }
            }
        }


        private void _CreateProgressBar()
        {
            //'Create a custom progress box ===========================================================
            try
            {
                SAPbouiCOM.Form oFormA  = _App.Forms.Item("frmCount");
                oFormA.Close();
            }
            catch{}
            
            oTimer = new System.Threading.Timer(new System.Threading.TimerCallback(TimerKeepAlive));
            oTimer.Change(0, 60 * 1000); //Timer to clear the windows message queue


            SAPbouiCOM.FormCreationParams oFormCP;
            oFormCP = (SAPbouiCOM.FormCreationParams)_App.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
            oFormCP.FormType = "frmCount";
            oFormCP.UniqueID = "frmCount";
            oFormCP.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_FixedNoTitle;
            _oForm = _App.Forms.AddEx(oFormCP);
            _FormUID = _oForm.UniqueID;
            _oForm.Height = 40;
            _oForm.Width = 250;
            _oForm.Left = _X;
            _oForm.Top = _Y;
            _oForm.Visible = true;

            if (_Type == ProgressType.ValueCount)
            {
                SAPbouiCOM.Item oItem1 = _oForm.Items.Add("lbl1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem1.Left = 10;
                oItem1.Top = 10;
                oItem1.Width = 135;
                SAPbouiCOM.StaticText oLabelCount = (SAPbouiCOM.StaticText)oItem1.Specific;
                oLabelCount.Caption = _Title;
                SAPbouiCOM.Item oItem = _oForm.Items.Add("txtCount", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                _oForm.DataSources.UserDataSources.Add("txtCount", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 8);
                oItem.Left = 150;
                oItem.Width = 40;
                oItem.Top = 10;
                oItem1.LinkTo = oItem.UniqueID;
                oItem.RightJustified = true;
                oItem.Enabled = false;
                SAPbouiCOM.EditText oTxt = (SAPbouiCOM.EditText)oItem.Specific;
                oTxt.DataBind.SetBound(true, "", "txtCount");
                oTxt.Value = "0";
                oItem1 = _oForm.Items.Add("lblCount", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem1.Left = oItem.Left + oItem.Width + 3;
                oItem1.Top = 10;
                oItem1.Width = 50;
                oLabelCount = (SAPbouiCOM.StaticText)oItem1.Specific;
                oLabelCount.Caption = "/ " + _MaxValue.ToString();
            }
            else if (_Type == ProgressType.ProgressBar)
            {
                SAPbouiCOM.Item oItem1 = _oForm.Items.Add("lbl1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem1.Left = 10;
                oItem1.Top = 3;
                oItem1.Width = 135;
                SAPbouiCOM.StaticText oLabelCount = (SAPbouiCOM.StaticText)oItem1.Specific;
                oLabelCount.Caption = String.Format("{0} {1} Records ", _Title, _MaxValue);
                SAPbouiCOM.Item oItem = _oForm.Items.Add("rct1", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
                oItem.Left = oItem1.Left;
                oItem.Top = oItem1.Top + oItem1.Height + 2;
                oItem.Width = _oForm.Width - 20;
                oItem.Height = 16;

                oItem1 = _oForm.Items.Add("lblBar", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem1.BackColor = 10638900;
                oItem1.Left = oItem.Left + 1;
                oItem1.Width = 0;
                oItem1.Top = oItem.Top + 1;

                _BarWidth = oItem.Width - 2;
            }
            else
            {
                SAPbouiCOM.Item oItem1 = _oForm.Items.Add("lbl1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem1.Left = 10;
                oItem1.Top = 3;
                oItem1.Width = 135;
                SAPbouiCOM.StaticText oLabelCount = (SAPbouiCOM.StaticText)oItem1.Specific;
                oLabelCount.Caption = _Title;
                SAPbouiCOM.Item oItem = _oForm.Items.Add("txtCount", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                _oForm.DataSources.UserDataSources.Add("txtCount", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 8);
                oItem.Left = 150;
                oItem.Width = 40;
                oItem.Top = 3;
                oItem.RightJustified = true;
                oItem1.LinkTo = oItem.UniqueID;
                oItem.Enabled = false;
                SAPbouiCOM.EditText oTxt = (SAPbouiCOM.EditText)oItem.Specific;
                oTxt.DataBind.SetBound(true, "", "txtCount");
                oTxt.Value = "0";
                oItem1 = _oForm.Items.Add("lblCount", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem1.Left = oItem.Left + oItem.Width + 3;
                oItem1.Top = 3;
                oItem1.Width = 50;
                oLabelCount = (SAPbouiCOM.StaticText)oItem1.Specific;
                oLabelCount.Caption = "/ " + _MaxValue.ToString();

                oItem = _oForm.Items.Add("rct1", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
                oItem.Left = 10;
                oItem.Top = oItem1.Top + oItem1.Height + 2;
                oItem.Width = _oForm.Width - 20;
                oItem.Height = 16;

                oItem1 = _oForm.Items.Add("lblBar", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem1.BackColor = 10638900;
                oItem1.Left = oItem.Left + 1;
                oItem1.Width = 0;
                oItem1.Top = oItem.Top + 1;

                _BarWidth = oItem.Width - 2;

            }


            //End Creating form =========================================================================

        }

        private void TimerKeepAlive(object State)
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
            if (oTimer != null)
                oTimer.Dispose();
            _oForm.Close();
            
        }

        #endregion

    }
}
