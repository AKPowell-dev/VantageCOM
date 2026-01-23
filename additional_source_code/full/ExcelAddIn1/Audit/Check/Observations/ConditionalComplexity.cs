using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class ConditionalComplexity : Observation
{
	public ConditionalComplexity(Severity sev, Range rng, int intIfs)
		: base(Category.FormulaComplexity, sev, VH.A(11736), rng)
	{
		base.Subtitle = intIfs + VH.A(11781) + rng.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		base.Explanation = VH.A(11818);
	}
}
