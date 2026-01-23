using A;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class LegacyFileType : Observation
{
	public LegacyFileType(Severity sev)
		: base(Category.Workbook, sev, VH.A(20039))
	{
		base.Subtitle = VH.A(6144);
		base.Explanation = VH.A(20072);
	}
}
