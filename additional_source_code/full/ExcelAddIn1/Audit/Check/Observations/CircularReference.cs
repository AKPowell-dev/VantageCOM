using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class CircularReference : Observation
{
	public CircularReference(Severity sev, Range rng)
		: base(Category.FormulaIntegrity, sev, VH.A(11251), rng)
	{
		base.Explanation = VH.A(11288);
	}
}
