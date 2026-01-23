using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class ExtraneousSheetName : Observation
{
	public ExtraneousSheetName(Severity sev, Range rng)
		: base(Category.FormulaErrors, sev, VH.A(16764), rng)
	{
		base.Explanation = VH.A(16807);
	}
}
