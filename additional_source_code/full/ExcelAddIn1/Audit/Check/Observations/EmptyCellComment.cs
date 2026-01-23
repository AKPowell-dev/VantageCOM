using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class EmptyCellComment : Observation
{
	public EmptyCellComment(Severity sev, Range rng)
		: base(Category.Oddities, sev, VH.A(14568), rng)
	{
		base.Explanation = VH.A(14605);
	}
}
