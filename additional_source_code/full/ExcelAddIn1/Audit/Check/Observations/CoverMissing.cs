using A;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class CoverMissing : Observation
{
	public CoverMissing(Severity sev)
		: base(Category.BestPractices, sev, VH.A(12521))
	{
		base.Subtitle = VH.A(12548);
		base.Explanation = VH.A(12631);
	}
}
