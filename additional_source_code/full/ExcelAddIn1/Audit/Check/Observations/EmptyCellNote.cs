using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class EmptyCellNote : Observation
{
	public EmptyCellNote(Severity sev, Range rng)
		: base(Category.Oddities, sev, VH.A(14869), rng)
	{
		base.Explanation = VH.A(14900);
	}
}
