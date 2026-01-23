using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class EmptyCellReference : Observation
{
	public EmptyCellReference(Severity sev, Range rng)
		: base(Category.FormulaErrors, sev, VH.A(15158), rng)
	{
		base.Explanation = VH.A(15199);
	}
}
