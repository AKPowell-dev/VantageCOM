using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class TooManyPrecedents : Observation
{
	public TooManyPrecedents(Severity sev, Range rng, int intPrecedents)
		: base(Category.FormulaComplexity, sev, VH.A(24228), rng)
	{
		base.Subtitle = intPrecedents + VH.A(24259) + rng.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		base.Explanation = VH.A(24288);
	}
}
