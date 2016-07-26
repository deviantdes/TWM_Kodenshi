using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TWM_KDS_AddOn
{
    [AttributeUsage(AttributeTargets.Class, Inherited = false, AllowMultiple = false)]
    sealed class AuthorizationAttribute : Attribute
    {
        public readonly string FormType = "";
        public readonly string Name = "";
        public readonly string ParentID = "";
        public readonly SAPbobsCOM.BoUPTOptions Options = SAPbobsCOM.BoUPTOptions.bou_FullReadNone;

        public AuthorizationAttribute(String formtype, String name, String parentID, SAPbobsCOM.BoUPTOptions option)
        {
            this.FormType = formtype;
            this.Name = name;
            this.ParentID = parentID;
            this.Options = option;
        }
    }
}
