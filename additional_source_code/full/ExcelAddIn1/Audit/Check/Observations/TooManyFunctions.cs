using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class TooManyFunctions : Observation
{
	public TooManyFunctions(Severity sev, Range rng, int intFunctions)
		: base(Category.FormulaComplexity, sev, VH.A(22908), rng)
	{
		base.Subtitle = intFunctions + VH.A(22937) + rng.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		base.Explanation = VH.A(22964);
	}
}
