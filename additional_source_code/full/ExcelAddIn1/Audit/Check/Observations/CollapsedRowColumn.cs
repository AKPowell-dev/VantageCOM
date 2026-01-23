using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class CollapsedRowColumn : Observation
{
	internal CollapsedRowColumn(Severity A, Worksheet B, Range C)
		: base(Category.HiddenData, A, VH.A(11486))
	{
		base.Sheet = B;
		base.Range = C;
		base.Subtitle = VH.A(11531) + C.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		base.Explanation = VH.A(11536);
	}
}
