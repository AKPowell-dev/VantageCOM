using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class FormulaTooLong : Observation
{
	public FormulaTooLong(Severity sev, Range rng)
		: base(Category.FormulaComplexity, sev, VH.A(17764), rng)
	{
		base.Explanation = VH.A(17789);
	}
}
