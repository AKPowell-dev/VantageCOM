using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class ConditionalFormat : Observation
{
	public ConditionalFormat(Severity sev, Range rng)
		: base(Category.Performance, sev, VH.A(12168), rng)
	{
		base.Explanation = VH.A(12213);
	}
}
