//****************************************************************************
//
//  File:      B1DbValidValue.cs
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

namespace B1WizardBase
{
  /// <summary>
  /// Class that represents a valid value inside a Field.
  /// </summary>
  /// <remarks>
  /// This class will be used by the class managing the Database (class inheriting 
  /// from B1Db base class).
  /// </remarks>
  public class B1DbValidValue
  {
    /// <summary>
    /// Value of the ValidValue.
    /// </summary>
    public string Val;
    /// <summary>
    /// Description of the ValidValue.
    /// </summary>
    public string Description;

    /////////////////////////////////////////////////////////////////////////////

    /// <summary>
    /// Empty Constructor.
    /// </summary>
    public B1DbValidValue()
    {
    }
 
    /// <summary>
    /// Creates a B1DbValidValue.
    /// </summary>
    /// <param name="val">Value of the ValidValue.</param>
    /// <param name="description">Description of the ValidValue.</param>
    public B1DbValidValue(
      string val,
      string description)
    {
      this.Val = val;
      this.Description = description;
    }
  }
}
