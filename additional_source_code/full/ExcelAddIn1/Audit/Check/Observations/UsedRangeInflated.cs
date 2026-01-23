using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class UsedRangeInflated : Observation
{
	public UsedRangeInflated(Severity sev, Range rngUsed)
		: base(Category.Performance, sev, VH.A(25547), rngUsed)
	{
		base.Subtitle = rngUsed.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(25588) + rngUsed.Worksheet.Name;
		base.Explanation = VH.A(25597);
	}
}
