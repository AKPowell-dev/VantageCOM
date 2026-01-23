using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class DeprecatedFunction : Observation
{
	public DeprecatedFunction(Severity sev, Range rng, string strFunction)
		: base(Category.FormulaIntegrity, sev, VH.A(13563), rng)
	{
		base.Subtitle = strFunction + VH.A(9959) + rng.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		base.Explanation = VH.A(13602);
	}
}
