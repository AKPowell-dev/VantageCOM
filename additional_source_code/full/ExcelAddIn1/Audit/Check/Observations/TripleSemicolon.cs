using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class TripleSemicolon : Observation
{
	public TripleSemicolon(Severity sev, Range rng)
		: base(Category.HiddenData, sev, VH.A(24660), rng)
	{
		base.Subtitle = rng.NumberFormatLocal.ToString() + VH.A(24693) + rng.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		base.Explanation = VH.A(24728);
	}
}
