//****************************************************************************
//
//  File:      B1DbKey.cs
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
  /*
  enum BoFieldTypes {
    db_Alpha = 0,
    db_Memo = 1,
    db_Numeric = 2,
    db_Date = 3,
    db_Float = 4
    };
  */

  /// <summary>
  /// Manages the B1 SDK metadata object UserKeysMD.
  /// </summary>
  /// <remarks>
  /// This class will be used by the class managing the Database (class inheriting 
  /// from B1Db base class).
  /// </remarks>
  public class B1DbKey
  {
    /// <summary>
    /// Name of the table this key refers to without "@".
    /// </summary>
    public string				TableName;	/// fields.item[0]

    /// <summary>
    /// Key name.
    /// </summary>
    public string				Name;				/// fields.item[1]
    
    /// <summary>
    /// Valid value of BoYesNoEnum type that specifies whether 
    /// or not the key is unique.
    /// </summary>
    public bool					IsUnique;			/// fields.item[2]

    /// <summary>
    /// List of fields that combine the key index (without "U_").
    /// </summary>
    public String[]	Elements = new string[0];  /// fields.item[3]

    /// <summary>
    /// Boolean value that shows if the key is primary or not.
    /// ReadOnly property, we cannot define a primary key for a UserTable.
    /// </summary>
    public bool         IsPrimary;

    /////////////////////////////////////////////////////////////////////////////

    /// <summary>
    /// Empty Constructor
    /// </summary>
    public B1DbKey()
    {
    }

    /// <summary>
    /// Builds a B1DbKey from its main members information.
    /// </summary>
    /// <param name="table">Name of the table this key refers to.</param>
    /// <param name="name">Key name.</param>
    /// <param name="isUnique">Valid value of BoYesNoEnum type that specifies whether or not the key is unique.</param>
    /// <param name="elements">List of fields that combine the key index.</param>
    public B1DbKey(
      string table,
      string name,
      bool isUnique ,
      string[] elements)
    {
      this.TableName = table;
      this.Name = name;
      this.IsUnique = isUnique;
      this.Elements = elements;
    }

    /// <summary>
    /// Adds the UserKeyMD to the current company Database.
    /// </summary>
    /// <param name="company">SAPbobsCOM.Company we are connected to.</param>
    /// <returns>Return value from the SDK action UserKeysMD.Add().</returns>
    public int Add(Company company)
    {
			UserKeysMD userKeys = null;
			int ret = -1;

			try
			{
				userKeys = (UserKeysMD)
					company.GetBusinessObject(BoObjectTypes.oUserKeys);
				userKeys.TableName = TableName;
				userKeys.KeyName = Name;
				userKeys.Unique = (IsUnique) ? BoYesNoEnum.tYES : BoYesNoEnum.tNO;

				//// elements
				bool flagFirst = true;
				foreach (string fieldName in Elements)
				{
					if (flagFirst)
						flagFirst = false;
					else
						userKeys.Elements.Add();
					userKeys.Elements.ColumnAlias = fieldName;
				}

				ret = userKeys.Add();

#if	DEBUG
				if (ret != 0)
				{
					int errcode;
					string errmsg;
					company.GetLastError(out errcode, out errmsg);
					System.Console.Out.WriteLine("Key " + Name + " : " + errmsg);
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(userKeys);
				userKeys = null;
			}
      return ret;
    }

    /// <summary>
    /// Generates the code to add a new UserKey. 
    /// <para>This code is added in your AddOn_Db class inheriting from B1Db.</para>
    /// </summary>
    /// <returns>CodeExpression containing the UserKey information.</returns>
    public CodeExpression GenerateCtor()
    {
      /*
        new B1DbKey(table,name,unique,elements);
      */

      int i = 0;
      CodeExpression[] elementsArray = new CodeExpression[ Elements.Length ];
      foreach	(string fieldName in Elements)
        elementsArray[ i++ ] = new CodePrimitiveExpression( fieldName );
      CodeArrayCreateExpression createElementsArray = new CodeArrayCreateExpression(
        "System.String", elementsArray);

      return new CodeObjectCreateExpression(
        "B1DbKey",
        new CodeExpression[4] {
                                new CodePrimitiveExpression( TableName ),
                                new CodePrimitiveExpression( Name ),
                                new CodePrimitiveExpression( IsUnique ),
                                createElementsArray});
    }

    /////////////////////////////////////////////////////////////////////////////

    /// <summary>
    /// Determines whether the specified B1DbKey is equal to the current B1DbKey.     
    /// </summary>
    /// <param name="obj">B1DbKey to compare.</param>
    /// <returns>true if both objects are equal.</returns>
    public override bool Equals(object obj)
    {
      if	(obj is B1DbKey)
      {
        B1DbKey key = (B1DbKey)obj;
        return (key.TableName.Equals(TableName))&&(key.Name.Equals(Name));
      }

      return base.Equals (obj);
    }

    /// <summary>
    /// Serves as a hash function for a particular type, suitable 
    /// for use in hashing algorithms and data structures like a hash table. 
    /// </summary>
    /// <returns>A hash code for the current B1DbKey.</returns>
    public override int GetHashCode()
    {
      return base.GetHashCode ();
    }

  }
}
