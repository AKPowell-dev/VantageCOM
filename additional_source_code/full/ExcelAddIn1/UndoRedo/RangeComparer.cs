using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.UndoRedo;

public sealed class RangeComparer : IEqualityComparer<Range>
{
	public bool Equals(Range x, Range y)
	{
		bool result = default(bool);
		try
		{
			if (x != null)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (y != null)
				{
					result = Operators.CompareString(x.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), y.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) == 0;
					return result;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
			}
			result = false;
			return result;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	bool IEqualityComparer<Range>.Equals(Range x, Range y)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(x, y);
	}

	public int GetHashCode(Range obj)
	{
		return obj.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)).GetHashCode();
	}

	int IEqualityComparer<Range>.GetHashCode(Range obj)
	{
		//ILSpy generated this explicit interface implementation from .override directive in GetHashCode
		return this.GetHashCode(obj);
	}
}
