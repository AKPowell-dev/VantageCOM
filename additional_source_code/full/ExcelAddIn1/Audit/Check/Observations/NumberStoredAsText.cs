using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class NumberStoredAsText : Observation
{
	public NumberStoredAsText(Severity sev, Range rng)
		: base(Category.Data, sev, VH.A(20608), rng)
	{
		if (Operators.ConditionalCompareObjectEqual(rng.Cells.CountLarge, 1, TextCompare: false))
		{
			base.Explanation = VH.A(20651);
		}
		else
		{
			base.Explanation = VH.A(20805);
		}
		base.HasFix = true;
		base.CanFixMultiple = true;
	}

	public override void FixAction()
	{
		base.FixAction();
		base.Range.NumberFormat = VH.A(20593);
		base.Range.Value2 = RuntimeHelpers.GetObjectValue(base.Range.Value2);
	}
}
