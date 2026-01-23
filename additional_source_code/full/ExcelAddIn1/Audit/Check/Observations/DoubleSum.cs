using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class DoubleSum : Observation
{
	public DoubleSum(Severity sev, Range rng)
		: base(Category.FormulaErrors, sev, VH.A(13986), rng)
	{
		base.Explanation = VH.A(14007);
	}
}
