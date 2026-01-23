using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class UnusedNumericInput : Observation
{
	public UnusedNumericInput(Severity sev, Range rng)
		: base(Category.FormulaErrors, sev, VH.A(25288), rng)
	{
		base.Explanation = VH.A(25329);
	}
}
