using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class FormulaError : Observation
{
	public FormulaError(Severity sev, Range rng)
		: base(Category.FormulaErrors, sev, VH.A(17323), rng)
	{
		if (Operators.ConditionalCompareObjectEqual(rng.CountLarge, 1, TextCompare: false))
		{
			while (true)
			{
				switch (2)
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
			base.Subtitle = rng.Text.ToString() + VH.A(17350) + rng.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		base.Explanation = VH.A(17357);
	}
}
