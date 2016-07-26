using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TWM_KDS_AddOn
{
    [AttributeUsage(AttributeTargets.Class, Inherited = false, AllowMultiple = false)]
    sealed class FormAttribute : Attribute
    {
        public readonly string FormType;
        public readonly bool HasMenu;
        public readonly string MenuName;
        public readonly string ParentMenu;
        public readonly int Position;
        public string TypeName { get; set; }
        
        public FormAttribute(string formtype)
        {
            this.FormType = formtype;
            this.HasMenu = false;
            this.MenuName = "";
            this.ParentMenu = "";
            this.Position = -1;
            this.TypeName = "";
        }

        public FormAttribute(string formtype, bool hasmenu, string menuname, string parentmenu, int pos)
        {
            this.FormType = formtype;
            this.HasMenu = hasmenu;
            this.MenuName = menuname;
            this.ParentMenu =parentmenu;
            this.Position = pos;
            this.TypeName = "";
        }
    }

}
