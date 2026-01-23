using A;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class HiddenName : Observation
{
	internal HiddenName(Severity A)
		: base(Category.HiddenData, A, VH.A(8786))
	{
		base.Subtitle = VH.A(8815);
		base.Explanation = VH.A(8900);
	}
}
