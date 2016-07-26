//****************************************************************************
//
//  File:      B1Udo.cs
//
//  Copyright (c) SAP 
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
//****************************************************************************
using System;
using System.CodeDom;
using SAPbobsCOM;

namespace B1WizardBase
{
  /// <summary>
  /// Manages the B1 SDK metadata object UserObjectsMD.
  /// </summary>
  /// <remarks>
  /// This class will be used by the class managing the Database (class inheriting 
  /// from B1Db base class).
  /// </remarks>
  public class B1Udo
  {
    /// <summary>
    /// Object Unique ID. 
    /// The Unique ID is the primary key of the user defined object.
    /// </summary>
    public string		    Code;
    /// <summary>
    /// Object name that must include your namespace.
    /// </summary>
    public string		    Name;
    /// <summary>
    /// Name of the main User Table related to the user defined object.
    /// </summary>
    public string		    Table;
    /// <summary>
    /// Collection of Child User Tables names.
    /// </summary>
    public string[]	    Children = new string[0];
    /// <summary>
    /// Valid value of BoUDOObjType type that specifies the object type:
    /// Master Data or Document. 
    /// </summary>
    public BoUDOObjType Type;
    /// <summary>
    /// Valid value of BoYesNoEnum type that specifies whether or not the 
    /// user defined object can use the Find service. 
    /// This service enables the Choose from list dialog box in the application (Find Mode). 
    /// </summary>
    public BoYesNoEnum  CanFind;
    /// <summary>
    /// Valid value of BoYesNoEnum type that specifies whether or not the user 
    /// defined object can use the Delete service. 
    /// This service enables to delete data from user tables of Master Data type objects. 
    /// </summary>
    public BoYesNoEnum  CanDelete;
    /// <summary>
    /// Valid value of BoYesNoEnum type that specifies whether or not the user
    /// defined object can use the Cancel service. 
    /// This service enables to cancel the user defined object.
    /// </summary>
    public BoYesNoEnum  CanCancel;
    /// <summary>
    /// valid value of BoYesNoEnum type that specifies whether or not the user 
    /// defined object can use the Close service. 
    /// This service enables to close a user defined object without creating 
    /// a posting in accounting.
    /// </summary>
    public BoYesNoEnum  CanClose;
    /// <summary>
    /// valid value of BoYesNoEnum type that specifies whether or not the user 
    /// defined object can use the Year Transfer service. 
    /// This service enables to copy the user tables to a new database.
    /// </summary>
    public BoYesNoEnum  CanYearTransfer;
    /// <summary>
    /// valid value of BoYesNoEnum type that specifies whether or not the user 
    /// defined object can use the History Log service. 
    /// This service creates a history log table in the database.
    /// </summary>
    public BoYesNoEnum  CanLog;
    /// <summary>
    /// valid value of BoYesNoEnum type that specifies whether or not the user 
    /// defined object can use the Manage Series service. 
    /// This service enables document numbering.
    /// </summary>
    public BoYesNoEnum  ManageSeries;
    /// <summary>
    /// Log table name. 
    /// This table maintains a history log of all actions related to the main user table.
    /// </summary>
    public string      LogTableName;
    /// <summary>
    /// Represents a collection of the fields alias to display in the 
    /// Find Form (Choose From List form).
    /// </summary>
    public string[]  FindColumnsAlias = new string[0];
    /// <summary>
    /// Represents a collection of the fields description to display 
    /// in the Find Form (Choose From List form).
    /// </summary>
    public string[]  FindColumnsDesc = new string[0];

    /////////////////////////////////////////////////////////////////////////////

    /// <summary>
    /// Empty Constructor.
    /// </summary>
    public B1Udo()
    {
    }

    /// <summary>
    /// Builds a B1Udo given its members information.
    /// </summary>
    /// <param name="code">Object Unique ID.</param>
    /// <param name="name">Object name that must include your namespace.</param>
    /// <param name="table">Name of the main User Table</param>
    /// <param name="children">Collection of Child User Tables names.</param>
    /// <param name="type">Valid value of BoUDOObjType type that specifies the object type.</param>
    /// <param name="canFind">Valid value of BoYesNoEnum type that specifies whether or not the user defined object can use the Find service.</param>
    /// <param name="canDelete">Valid value of BoYesNoEnum type that specifies whether or not the user defined object can use the Delete service.</param>
    /// <param name="canCancel">Valid value of BoYesNoEnum type that specifies whether or not the user defined object can use the Cancel service.</param>
    /// <param name="canClose">Valid value of BoYesNoEnum type that specifies whether or not the user defined object can use the Close service.</param>
    /// <param name="canYearTransfer">Valid value of BoYesNoEnum type that specifies whether or not the user defined object can use the Year Transfer service.</param>
    /// <param name="canLog">Valid value of BoYesNoEnum type that specifies whether or not the user defined object can use the History Log service.</param>
    /// <param name="manageSeries">Valid value of BoYesNoEnum type that specifies whether or not the user defined object can use the Manage Series service.</param>
    /// <param name="logTableName">Log table name.</param>
    /// <param name="findColumnsAlias">Collection of the fields alias to display in the Find Form.</param>
    /// <param name="findColumnsDesc">Collection of the fields description to display in the Find Form.</param>
    public B1Udo(
      string code,
      string name,
      string table,
      string[] children,
      BoUDOObjType type,
      BoYesNoEnum canFind,
      BoYesNoEnum canDelete,
      BoYesNoEnum canCancel,
      BoYesNoEnum canClose,
      BoYesNoEnum canYearTransfer,
      BoYesNoEnum canLog,
      BoYesNoEnum manageSeries,
      string logTableName,
      string[] findColumnsAlias,
      string[] findColumnsDesc)
    {
      this.Code = code;
      this.Name = name;
      this.Table = table;
      this.Children = children;
      this.Type = type;
      this.CanFind = canFind;
      this.CanDelete = canDelete;
      this.CanCancel = canCancel;
      this.CanClose = canClose;
      this.ManageSeries = manageSeries;
      this.CanYearTransfer = canYearTransfer;
      this.CanLog = canLog;
      this.LogTableName = logTableName;
      this.FindColumnsAlias = findColumnsAlias;
      this.FindColumnsDesc = findColumnsDesc;
    }

    /// <summary>
    /// Adds the UDO object into the current company Database.
    /// </summary>
    /// <param name="company">SAPbobsCOM.Company we are connected to.</param>
    /// <returns>Return value of the UserObjectsMD.Add() call.</returns>
    public int Add(Company company)
    {
      UserObjectsMD udo = (UserObjectsMD)
        company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
      udo.Code = Code;
      udo.Name = Name;
      udo.TableName = Table;
      udo.ObjectType = Type; 

      foreach	(string child in Children)
      {
        udo.ChildTables.TableName = child;
        udo.ChildTables.Add();
      }

      // Services Definition
      udo.CanFind = CanFind;
      if (CanFind == BoYesNoEnum.tYES)
      {
        for (int i = 0; i < FindColumnsAlias.GetLength(0); i++)
        {
          udo.FindColumns.ColumnAlias = FindColumnsAlias[i];
          udo.FindColumns.ColumnDescription = FindColumnsDesc[i];
          udo.FindColumns.Add();
        }
      }
      udo.CanDelete = CanDelete;
      udo.CanCancel = CanCancel;
      udo.CanClose = CanClose;
      udo.ManageSeries = ManageSeries;
      udo.CanYearTransfer = CanYearTransfer;
      udo.CanLog = CanLog;
      udo.LogTableName = LogTableName;

      int ret = udo.Add();

#if	DEBUG
			if	(ret != 0)
			{
				int errcode;
				string errmsg;
				company.GetLastError(out errcode, out errmsg);
				System.Console.Out.WriteLine("UDO " + Name + " : " + errmsg);
			}
#endif

      // clean DI object
			System.Runtime.InteropServices.Marshal.ReleaseComObject(udo);
			udo = null;
      //System.GC.Collect();
      //System.GC.WaitForPendingFinalizers();

      return ret;
    }

    /// <summary>
    /// Generates the code to add a new UserObject (UDO). 
    /// <para>This code is added in your AddOn_Db class inheriting from B1Db.</para>
    /// </summary>
    /// <returns>CodeExpression containing the UserObject information.</returns>
    public CodeExpression GenerateCtor()
    {
      /*
        new B1Udo(code,name,table,children,type, 
                  canFind, canDelete, canCancel, 
                  canClose, canYearTransfer, canLog, manageSeries);
      */

      int i = 0;
      CodeExpression[] childrenArray = new CodeExpression[ Children.Length ];
      foreach	(string child in Children)
        childrenArray[ i++ ] = new CodePrimitiveExpression( child );
      CodeArrayCreateExpression createChildrenArray = new CodeArrayCreateExpression(
        "System.String",childrenArray);
		
      i = 0;
      CodeExpression[] findAliasArray = new CodeExpression[ FindColumnsAlias.Length ];
      foreach	(string column in FindColumnsAlias)
        findAliasArray[ i++ ] = new CodePrimitiveExpression( column );
      CodeArrayCreateExpression createFindAliasArray = new CodeArrayCreateExpression(
        "System.String", findAliasArray);

      i = 0;
      CodeExpression[] findDescArray = new CodeExpression[ FindColumnsDesc.Length ];
      foreach	(string column in FindColumnsDesc)
        findDescArray[ i++ ] = new CodePrimitiveExpression( column );
      CodeArrayCreateExpression createFindDescArray = new CodeArrayCreateExpression(
        "System.String", findDescArray);

      //-
      return new CodeObjectCreateExpression(
        "B1Udo",
        new CodeExpression[15] {
                                 new CodePrimitiveExpression( Code ),
                                 new CodePrimitiveExpression( Name ),
                                 new CodePrimitiveExpression( Table ),
                                 createChildrenArray,
                                 new CodeFieldReferenceExpression(
                                 new CodeTypeReferenceExpression("BoUDOObjType"),Type.ToString()),
                                 new CodeFieldReferenceExpression(
                                 new CodeTypeReferenceExpression("BoYesNoEnum"), CanFind.ToString()),
                                 new CodeFieldReferenceExpression(
                                 new CodeTypeReferenceExpression("BoYesNoEnum"), CanDelete.ToString()),
                                 new CodeFieldReferenceExpression(
                                 new CodeTypeReferenceExpression("BoYesNoEnum"), CanCancel.ToString()),
                                 new CodeFieldReferenceExpression(
                                 new CodeTypeReferenceExpression("BoYesNoEnum"), CanClose.ToString()),
                                 new CodeFieldReferenceExpression(
                                 new CodeTypeReferenceExpression("BoYesNoEnum"), CanYearTransfer.ToString()),
                                 new CodeFieldReferenceExpression(
                                 new CodeTypeReferenceExpression("BoYesNoEnum"), CanLog.ToString()),
                                 new CodeFieldReferenceExpression(
                                 new CodeTypeReferenceExpression("BoYesNoEnum"), ManageSeries.ToString()),
                                 new CodePrimitiveExpression( LogTableName ),
                                 createFindAliasArray,
                                 createFindDescArray});
    }

    /////////////////////////////////////////////////////////////////////////////

    /// <summary>
    /// Determines whether the specified B1DbTable is equal to the current B1Udo.     
    /// </summary>
    /// <param name="obj">B1Udo to compare.</param>
    /// <returns>True if both objects are equal.</returns>
    public override bool Equals(object obj)
    {
      if	(obj is B1Udo)
      {
        B1Udo udo = (B1Udo)obj;
        return udo.Code.Equals(Code);
      }

      return base.Equals (obj);
    }

    /// <summary>
    /// Serves as a hash function for a particular type, suitable 
    /// for use in hashing algorithms and data structures like a hash table. 
    /// </summary>
    /// <returns>A hash code for the current B1Udo.</returns>
    public override int GetHashCode()
    {
      return base.GetHashCode ();
    }

  }
}
