using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class ValidationFailed : Observation
{
	public ValidationFailed(Severity sev, Range rng)
		: base(Category.Data, sev, VH.A(25955), rng)
	{
		if (Operators.ConditionalCompareObjectEqual(rng.Cells.CountLarge, 1, TextCompare: false))
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					base.Explanation = VH.A(25990);
					return;
				}
			}
		}
		base.Explanation = VH.A(26276);
	}
}
