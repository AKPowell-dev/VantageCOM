using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class PartialInput : Observation
{
	public PartialInput(Severity sev, Range rng)
		: base(Category.FormulaIntegrity, sev, VH.A(22240), rng)
	{
		base.Explanation = VH.A(22267);
	}
}
