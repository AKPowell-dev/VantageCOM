using System.Collections.Generic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.TraceDialogs.Precedents;

public sealed class PrecedentComparer : IEqualityComparer<wpfPrecedents.Precedent>
{
	public bool Equals1(wpfPrecedents.Precedent x, wpfPrecedents.Precedent y)
	{
		if (x == y)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return true;
				}
			}
		}
		if (x != null)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			if (y != null)
			{
				return Operators.CompareString(x.Address, y.Address, TextCompare: false) == 0;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		return false;
	}

	bool IEqualityComparer<wpfPrecedents.Precedent>.Equals(wpfPrecedents.Precedent x, wpfPrecedents.Precedent y)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals1
		return this.Equals1(x, y);
	}

	public int GetHashCode1(wpfPrecedents.Precedent precedent)
	{
		if (precedent == null)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return 0;
				}
			}
		}
		int num = ((precedent.Value != null) ? precedent.Value.GetHashCode() : 0);
		int hashCode = precedent.Address.GetHashCode();
		return num ^ hashCode;
	}

	int IEqualityComparer<wpfPrecedents.Precedent>.GetHashCode(wpfPrecedents.Precedent precedent)
	{
		//ILSpy generated this explicit interface implementation from .override directive in GetHashCode1
		return this.GetHashCode1(precedent);
	}
}
