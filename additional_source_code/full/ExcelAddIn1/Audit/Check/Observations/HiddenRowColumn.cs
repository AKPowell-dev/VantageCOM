using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class HiddenRowColumn : Observation
{
	internal HiddenRowColumn(Severity A, Worksheet B, Range C)
		: base(Category.HiddenData, A, VH.A(18412))
	{
		base.Sheet = B;
		base.Range = C;
		base.Subtitle = VH.A(18451);
		base.Explanation = VH.A(18544);
	}
}
