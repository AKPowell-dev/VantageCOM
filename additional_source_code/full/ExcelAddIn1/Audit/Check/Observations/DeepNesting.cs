using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class DeepNesting : Observation
{
	public DeepNesting(Severity sev, Range rng, int intLevels)
		: base(Category.FormulaComplexity, sev, VH.A(13051), rng)
	{
		base.Subtitle = intLevels + VH.A(13076) + rng.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		base.Explanation = VH.A(13113);
	}
}
