using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class VolatileFunction : Observation
{
	public VolatileFunction(Severity sev, Range rng, string strFunction)
		: base(Category.Performance, sev, VH.A(27037), rng)
	{
		base.Subtitle = strFunction + VH.A(9959) + rng.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		base.Explanation = VH.A(27072);
	}
}
