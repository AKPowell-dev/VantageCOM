using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class OmittedReference : Observation
{
	public OmittedReference(Severity sev, Range rng)
		: base(Category.FormulaErrors, sev, VH.A(22059), rng)
	{
		base.Explanation = VH.A(22094);
	}
}
