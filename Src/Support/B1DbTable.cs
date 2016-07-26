//****************************************************************************
//
//  File:      B1DbTable.cs
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
  /// Manages the B1 SDK metadata object UserTablesMD.
  /// </summary>
  /// <remarks>
  /// This class will be used by the class managing the Database (class inheriting 
  /// from B1Db base class).
  /// </remarks>
  public class B1DbTable
  {
    /// <summary>
    /// Name for the user defined table.
    /// </summary>
    public string			Name;

    /// <summary>
    /// A string that describes the name and functionality of the table.
    /// </summary>
    public string			Description;

    /// <summary>
    /// Valid value of BoUTBTableType type that specifies the type of the user table.
    /// </summary>
    public BoUTBTableType	Type = BoUTBTableType.bott_NoObject;

    /////////////////////////////////////////////////////////////////////////////

    /// <summary>
    /// Empty Constructor.
    /// </summary>
    public B1DbTable()
    {
    }

    /// <summary>
    /// Builds a B1DbTable from its main members information.
    /// </summary>
    /// <param name="name">Name for the user defined table.</param>
    /// <param name="description">A string that describes the name and functionality of the table.</param>
    /// <param name="type">Valid value of BoUTBTableType type that specifies the type of the user table.</param>
    public B1DbTable(
      string name,
      string description,
      BoUTBTableType type)
    {
      this.Name = name;
      this.Description = description;
      this.Type = type;
    }

    /// <summary>
    /// Adds the UserTable to the current company Database.
    /// </summary>
    /// <param name="company">SAPbobsCOM.Company we are connected to.</param>
    /// <returns>Return value from the SDK action UserTablesMD.Add().</returns>
    public int Add(Company company)
    {
			UserTablesMD userTables = null;
			int ret = -1;

      //System.GC.Collect();
      //System.GC.WaitForPendingFinalizers();

			try
			{
				userTables = (UserTablesMD)company.GetBusinessObject(BoObjectTypes.oUserTables);
				userTables.TableName = Name.Substring(1);	// remove @
				userTables.TableType = Type;
				userTables.TableDescription = Description;

				ret = userTables.Add();

#if	DEBUG
				if (ret != 0)
				{
					int errcode;
					string errmsg;
					company.GetLastError(out errcode, out errmsg);
					System.Console.Out.WriteLine("Table " + Name + " : " + errmsg);
				}
#endif
			}
			catch (Exception ex) 
			{
				throw ex;
			}
			finally
			{
				// clean DI object
				System.Runtime.InteropServices.Marshal.ReleaseComObject(userTables);
				userTables = null;
				//System.GC.Collect();
				//System.GC.WaitForPendingFinalizers();
			}
      return ret;
    }

    /// <summary>
    /// Generates the code to add a new UserTable. 
    /// <para>This code is added in your AddOn_Db class inheriting from B1Db.</para>
    /// </summary>
    /// <returns>CodeExpression containing the UserTable information.</returns>
    public CodeExpression GenerateCtor()
    {
      /*
        new B1DbTable(name,description,type);
      */

      return new CodeObjectCreateExpression(
        "B1DbTable",
        new CodeExpression[3] {
                                new CodePrimitiveExpression( Name ),
                                new CodePrimitiveExpression( Description ),
                                new CodeFieldReferenceExpression(
                                new CodeTypeReferenceExpression("BoUTBTableType"),Type.ToString())});
    }

    /////////////////////////////////////////////////////////////////////////////

    /// <summary>
    /// Determines whether the specified B1DbTable is equal to the current B1DbTable.     
    /// </summary>
    /// <param name="obj">B1DbTable to compare.</param>
    /// <returns>true if both objects are equal.</returns>
    public override bool Equals(object obj)
    {
      if	(obj is B1DbTable)
      {
        B1DbTable table = obj as B1DbTable;
        return table.Name.Equals(Name);
      }

      return base.Equals (obj);
    }

    /// <summary>
    /// Serves as a hash function for a particular type, suitable 
    /// for use in hashing algorithms and data structures like a hash table. 
    /// </summary>
    /// <returns>A hash code for the current B1DbTable.</returns>
    public override int GetHashCode()
    {
      return base.GetHashCode ();
    }
  }
}
