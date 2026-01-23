using A;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class HiddenObjects : Observation
{
	internal HiddenObjects(Severity A)
		: base(Category.HiddenData, A, VH.A(18099))
	{
		base.Subtitle = VH.A(18128);
		base.Explanation = VH.A(18268);
	}
}
