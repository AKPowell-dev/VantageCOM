using System.Collections.Generic;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.UndoRedo;

public sealed class WorksheetComparer : IEqualityComparer<Worksheet>
{
	public bool Equals(Worksheet x, Worksheet y)
	{
		if (Operators.CompareString(x.Name, y.Name, TextCompare: false) == 0)
		{
			while (true)
			{
				switch (5)
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
			if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(x.Parent, null, VH.A(99969), new object[0], null, null, null), NewLateBinding.LateGet(y.Parent, null, VH.A(99969), new object[0], null, null, null), TextCompare: false))
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						return true;
					}
				}
			}
		}
		return false;
	}

	bool IEqualityComparer<Worksheet>.Equals(Worksheet x, Worksheet y)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(x, y);
	}

	public int GetHashCode(Worksheet obj)
	{
		return NewLateBinding.LateGet(obj.Cells[1, 1], null, VH.A(5814), new object[1] { true }, new string[1] { VH.A(68999) }, null, null).GetHashCode();
	}

	int IEqualityComparer<Worksheet>.GetHashCode(Worksheet obj)
	{
		//ILSpy generated this explicit interface implementation from .override directive in GetHashCode
		return this.GetHashCode(obj);
	}
}
