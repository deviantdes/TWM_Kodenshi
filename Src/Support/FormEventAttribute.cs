using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TWM_KDS_AddOn
{
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    sealed class FormEventAttribute : Attribute
    {
        public readonly Object oEventType;
        public readonly bool BeforeAction;

        public FormEventAttribute(Object EventType, bool Before)
        {
            this.oEventType = EventType;
            this.BeforeAction = Before;
            
        }

    }
}
