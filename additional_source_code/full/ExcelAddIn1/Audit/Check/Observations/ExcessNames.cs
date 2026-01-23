using A;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class ExcessNames : Observation
{
	public ExcessNames(Severity sev, int intCount)
		: base(Category.Performance, sev, VH.A(15357))
	{
		base.Subtitle = intCount + VH.A(15386);
		base.Explanation = VH.A(15439);
	}
}
