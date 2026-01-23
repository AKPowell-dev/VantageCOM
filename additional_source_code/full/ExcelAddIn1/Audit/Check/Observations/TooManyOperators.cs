using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class TooManyOperators : Observation
{
	public TooManyOperators(Severity sev, Range rng, int intOperators)
		: base(Category.FormulaComplexity, sev, VH.A(23798), rng)
	{
		base.Subtitle = intOperators + VH.A(23827) + rng.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		base.Explanation = VH.A(23854);
	}
}
