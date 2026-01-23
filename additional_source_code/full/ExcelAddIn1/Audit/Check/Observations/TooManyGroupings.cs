using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class TooManyGroupings : Observation
{
	public TooManyGroupings(Severity sev, Range rng, int intGroupings)
		: base(Category.FormulaComplexity, sev, VH.A(23304), rng)
	{
		base.Subtitle = intGroupings + VH.A(23333) + rng.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		base.Explanation = VH.A(23360);
	}
}
