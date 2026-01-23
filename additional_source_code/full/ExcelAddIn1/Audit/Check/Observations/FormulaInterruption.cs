using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class FormulaInterruption : Observation
{
	public FormulaInterruption(Severity sev, Range rng)
		: base(Category.FormulaErrors, sev, VH.A(17573), rng)
	{
		base.Explanation = VH.A(17614);
	}
}
