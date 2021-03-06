﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.34209
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace TWM_KDS_AddOn.Src.Resource {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Queries {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Queries() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("TWM_KDS_AddOn.Src.Resource.Queries", typeof(Queries).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to SELECT 
        ///T0.DocNum as [Doc No.],
        ///T1.DocEntry as [DocEntry],
        ///T1.LineNum as [LineNum],
        ///T1.ItemCode as [Item Code],
        ///T1.SubCatNum as [BP Catalog No.],
        ///T1.Dscription as [Description],
        ///T1.U_TWM_CUSTPONO as [Cust.PO No.],
        ///T1.U_TWM_CTBOXNO as [Carton Box No.],
        ///T1.U_TWM_ITEMCD as [Part No.]
        ///FROM OINV T0 JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry
        ///WHERE T1.DocEntry =&apos;{0}&apos;.
        /// </summary>
        internal static string TWM_GET_INV {
            get {
                return ResourceManager.GetString("TWM_GET_INV", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to SELECT &apos;{1}&apos; Checked
        ///	, T1.ItemCode ItemNo
        ///	, T1.Dscription [Item Description]
        ///	, T1.OpenQty [Quantity]
        ///	, T1.Price [Unit Price]
        ///	, T1.OpenQty * T1.Price [Total Amount]
        ///	, T1.VatGroup [Tax Code]
        ///	, T1.U_TWM_CUSTPONO [PO Customer]
        ///	, T1.WhsCode [Warehouse]
        ///	, T1.ShipDate [DeliveryDate]
        ///	, T0.ObjType [Doc_Type]
        ///	, T0.DocEntry [Doc_Entry]
        ///	, T1.LineNum [Line_Num]
        ///FROM OPOR T0 JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry
        ///WHERE T0.DocEntry in ({0}) AND T1.OpenQty &gt; 0.
        /// </summary>
        internal static string TWM_GET_PO_GRID {
            get {
                return ResourceManager.GetString("TWM_GET_PO_GRID", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to SELECT ROW_NUMBER() OVER(ORDER BY CardCode) AS Row ,&apos;N&apos; as [Check], CardCode, CardName, DocNum, NumAtCard, DocEntry from OPOR where CardCode = &apos;{0}&apos; AND DocDate between &apos;{1}&apos; AND &apos;{2}&apos; AND DocStatus = &apos;O&apos;.
        /// </summary>
        internal static string TWM_GET_PO_MATRIX {
            get {
                return ResourceManager.GetString("TWM_GET_PO_MATRIX", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to SELECT &apos;{1}&apos; Checked
        ///	, T1.ItemCode ItemNo
        ///	, T1.SubCatNum [BP Catalog No]
        ///	, T1.Dscription [Item Description]
        ///	, T1.OpenQty [Quantity]
        ///	, T1.Price [Unit Price]
        ///	, T1.OpenQty * T1.Price [Total Amount]
        ///	, T1.U_TWM_CTBOXNO [Total Carton Box]
        ///	, T1.WhsCode [Warehouse]
        ///	, T1.ShipDate [Delivery Date]
        ///	, T0.ObjType [Doc_Type]
        ///	, T0.DocEntry [Doc_Entry]
        ///	, T1.LineNum [Line_Num]
        ///FROM ORDR T0 JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry
        ///WHERE T0.DocEntry in ({0}) AND T1.OpenQty &gt; 0.
        /// </summary>
        internal static string TWM_GET_SO_GRID {
            get {
                return ResourceManager.GetString("TWM_GET_SO_GRID", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to SELECT ROW_NUMBER() OVER(ORDER BY CardCode) AS Row ,&apos;N&apos; as [Check], CardCode, CardName, DocNum, NumAtCard, DocEntry from ORDR where CardCode = &apos;{0}&apos; AND DocDate between &apos;{1}&apos; AND &apos;{2}&apos; AND DocStatus = &apos;O&apos;.
        /// </summary>
        internal static string TWM_GET_SO_MATRIX {
            get {
                return ResourceManager.GetString("TWM_GET_SO_MATRIX", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to UPDATE INV1
        ///    SET U_TWM_CTBOXNO = CASE T1.U_TWM_CUSTPONO
        ///        {0}
        ///    END
        ///FROM OINV T0 JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry
        ///WHERE T1.U_TWM_CUSTPONO IN ({1}) AND T0.DocNum = {2}.
        /// </summary>
        internal static string TWM_UPDATE_DB {
            get {
                return ResourceManager.GetString("TWM_UPDATE_DB", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to UPDATE INV1
        ///    SET U_TWM_CTBOXNO = CASE 
        ///        {0}
        ///    END
        ///FROM INV1 
        ///WHERE  DocEntry IN ({1}).
        /// </summary>
        internal static string TWM_UPDATE_DB2 {
            get {
                return ResourceManager.GetString("TWM_UPDATE_DB2", resourceCulture);
            }
        }
    }
}
