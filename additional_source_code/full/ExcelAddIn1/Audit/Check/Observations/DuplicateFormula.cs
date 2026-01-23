using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class DuplicateFormula : Observation
{
	public DuplicateFormula(Severity sev, Range rng)
		: base(Category.FormulaIntegrity, sev, VH.A(14211), rng)
	{
		base.Explanation = VH.A(14246);
	}
}
