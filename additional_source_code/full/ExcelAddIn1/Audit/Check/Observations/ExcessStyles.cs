using A;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class ExcessStyles : Observation
{
	public ExcessStyles(Severity sev, int intCount)
		: base(Category.Performance, sev, VH.A(15945))
	{
		base.Subtitle = intCount + VH.A(15976);
		base.Explanation = VH.A(16031);
	}
}
