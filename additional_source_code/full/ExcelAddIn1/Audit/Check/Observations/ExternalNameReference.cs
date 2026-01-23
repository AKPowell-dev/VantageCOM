using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class ExternalNameReference : Observation
{
	public ExternalNameReference(Severity sev, Name nm)
		: base(Category.Workbook, sev, VH.A(16369))
	{
		base.Subtitle = VH.A(16412) + nm.Name + VH.A(16425);
		base.Explanation = VH.A(16470);
	}
}
